# Warehouse Movement App — Final Specification

- **Version:** `v0.5.15`
- **Repository:** https://github.com/Alaa-fa-hash/warehouse-movement-app
- **Release page:** https://github.com/Alaa-fa-hash/warehouse-movement-app/releases/tag/v0.5.15
- **Direct APK:** https://github.com/Alaa-fa-hash/warehouse-movement-app/releases/download/v0.5.15/app-release.apk
- **Source commit:** `5916780` (`Smooth PNG text rendering`)
- **License/status:** Internal application source snapshot as published in the repository.

## 1. Platform and reproducible build

The application is Android Native only: Kotlin with Jetpack Compose and Material 3. Expo, React Native, EAS, and external dashboards are not used.

Build configuration:

- JDK: 17
- Gradle Wrapper: 8.9
- Android Gradle Plugin: 8.7.2
- Kotlin: 1.9.24
- Compile SDK: 35
- Target SDK: 35
- Minimum SDK: 26
- Compose compiler extension: 1.5.14
- Compose BOM: `androidx.compose:compose-bom:2024.12.01`
- POI: `com.github.vince688:poi-android:3.17`
- CI: GitHub Actions on `ubuntu-latest`
- Release build command: `./gradlew :app:assembleRelease --no-daemon --stacktrace`
- Release workflow: `.github/workflows/android.yml`

The workflow creates the CI signing keystore, builds the release APK, and attaches `app-release.apk` to a GitHub Release for tags beginning with `v`.

## 2. Input Excel format

The primary supported format is a flat Excel table. Header discovery is based on header names rather than fixed column positions. The expected logical columns are:

`التاريخ | رمز المادة | اسم المادة | الكمية | الوحدة | نوع الفاتورة | مستودع القلم`

`مستودع القلم` is accepted as the salesperson/warehouse column. The parser also retains compatibility with the older pivot-style workbook through its stateful pivot parser; the flat-table path must not be validated against the old 12-column invoice schema.

For flat rows, quantities are read as positive magnitudes where appropriate. The movement type and salesperson are normalized from the header-mapped cells. The last imported workbook is copied into app storage and restored automatically after the app is closed and reopened.

## 3. Movement and calculation rules

### Wholesale (`جملة`)

Only transfer movements are used:

- `إدخال` / transfer input = loading.
- `إخراج` / transfer output = returned quantity.
- Sales invoice rows are ignored for the wholesale calculation.
- Actual sale is always calculated, never read from a precomputed Excel column:

`المبيع = التحميل - المرتجع`

Rows are grouped by salesperson, date, item code, item alias/name, and unit. The report columns are:

`التاريخ | المادة | الوحدة | تحميل | مرتجع | مبيع`

Zero rows are hidden when loading, returned, and sale are all zero. The total row is placed immediately after the last data row and its three numeric values are aligned below loading, returned, and sale.

### Retail (`مفرق`)

Only invoice rows classified as `مبيع` are used. Transfer input/output rows are ignored. The detailed retail report aggregates the same item into one row per date and displays unit quantities in separate columns:

`التاريخ | المادة | طرد | علبة`

The item is not duplicated when it has both carton and box sales. Rows with both carton and box equal to zero are hidden.

## 4. Priority and sorting rules

The user selects up to four priority materials by item code. The selection is persisted in `SharedPreferences` under `priority_codes` and is reused after restarting the app.

The priority order is exact:

1. selected material in slot 1
2. selected material in slot 2
3. selected material in slot 3
4. selected material in slot 4
5. all remaining materials

After the four priority materials, remaining rows are sorted descending by the sale quantity:

- Wholesale: `loading - returned`.
- Retail: `carton + box`.

This ordering is applied to the in-app report, PNG export, and XLSX export. The priority picker displays the saved/custom alias where available, followed by the item code, so the user selects the actual visible material name.

## 5. Material aliases

Aliases are persisted by item code in `SharedPreferences`. A generated alias is the first useful word from the long item name after removing digits, punctuation, and generic words.

The ignored generic words include:

- `شوكولا`
- `بسكويت`
- `بسكوت`
- `الصوص`
- `صوص`
- `ديانا`
- `رقم`
- `مادة`

The user can edit an alias from the material directory. A saved alias takes precedence in the app, priority picker, reports, PNG, and Excel.

## 6. In-app report design

The report is RTL and uses a clean white card with a pale header area, a turquoise table header, alternating neutral rows, and bold readable values.

- Row backgrounds alternate between `#FFFFFF` and `#F4F7F8`.
- The table header is turquoise/teal.
- The date column has additional width.
- Date and row values use `maxLines = 1` and `softWrap = false` so dates such as `24/11/2026` remain on one horizontal line.
- Report rows use a larger bold display size (`13sp`).
- The row height remains consistent and the date does not create a second line.

## 7. PNG export specification

PNG is lossless and is rendered as an Android `Bitmap.Config.ARGB_8888` bitmap, then compressed with `Bitmap.CompressFormat.PNG` at quality `100`.

- Logical phone width: `412px`.
- Pixel ratio: `4`.
- Actual output width: `1648px`.
- Height: automatic based on the number of visible rows.
- Direction: RTL layout/order.
- Font: Android `sans-serif` Bold, matching the in-app display font.
- Rendering: anti-aliasing, subpixel text, and linear text rendering enabled; artificial `FakeBold` is disabled to avoid rough or muddy glyphs.
- Header: clean pale background, salesperson name, date, and bold text.
- Summary: horizontal loading/returned/sale summary for wholesale, or retail sale summary.
- Table: clear turquoise header, alternating white/very-light neutral rows, bold dark non-zero values, and lightly muted zeros.
- Totals: directly below the final data row under the correct numeric columns.
- Footer: none. The image ends at the total row.
- Priority and sales sorting are applied before rendering.

The exported PNG is shared from the app cache through Android `FileProvider`. Sending it as a document/file preserves the original PNG; messaging apps may recompress it when sent as a normal image.

## 8. Excel/XLSX export specification

The export creates an XLSX workbook with one worksheet per salesperson. It uses `XSSFWorkbook` and does not call `autoSizeColumn()` because Android POI can raise `NoClassDefFoundError` for `FontRenderContext`.

- Fixed widths are applied with `setColumnWidth`.
- Worksheets use RTL via `setRightToLeft(true)`.
- Header fill: clear turquoise/teal, approximately `#0F99A8`.
- Header font: bold white.
- Alternating body rows: pure white `#FFFFFF` and neutral very-light gray `#F4F7F8`.
- Body and total cells have light thin borders and centered alignment.
- Total row is directly adjacent to the last data row.
- Wholesale totals align under `تحميل`, `مرتجع`, and `مبيع`.
- Retail totals align under `طرد` and `علبة`.
- Custom RGB fills are created through POI-compatible `CTColor`/`XSSFColor` APIs.
- The workbook is shared as an Open XML XLSX document through `FileProvider`.

## 9. Temporary storage and sharing

Generated PNG and XLSX files are written to the app cache under the `reports` directory and shared with Android's chooser using the configured `FileProvider` and `res/xml/file_paths.xml`. They are temporary exports and are not treated as permanent application data.

## 10. Archive contents and excluded files

This archive is generated from the exact `v0.5.15` Git tag. It includes source code, Gradle wrapper/configuration, Android resources, fonts, themes, manifest, Excel parser/exporter, GitHub Actions workflow, and project documentation.

Intentionally excluded for cleanliness and reproducible rebuilding:

- `.git/` history and internal Git metadata
- `.gradle/` caches
- `build/` and `app/build/` generated outputs
- IDE caches and temporary files
- local machine credentials and generated signing keystores

The signed APK remains available from the GitHub Release link at the top of this document.
