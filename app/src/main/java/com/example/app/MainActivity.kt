package com.example.app

import android.content.Context
import android.os.Bundle
import androidx.activity.ComponentActivity
import androidx.activity.compose.rememberLauncherForActivityResult
import androidx.activity.compose.setContent
import androidx.activity.result.contract.ActivityResultContracts
import androidx.compose.foundation.background
import androidx.compose.foundation.layout.Box
import androidx.compose.foundation.layout.Arrangement
import androidx.compose.foundation.layout.Column
import androidx.compose.foundation.layout.Row
import androidx.compose.foundation.layout.Spacer
import androidx.compose.foundation.layout.fillMaxSize
import androidx.compose.foundation.layout.fillMaxWidth
import androidx.compose.foundation.layout.height
import androidx.compose.foundation.layout.padding
import androidx.compose.foundation.lazy.LazyColumn
import androidx.compose.foundation.lazy.items
import androidx.compose.foundation.shape.RoundedCornerShape
import androidx.compose.material3.AlertDialog
import androidx.compose.material3.Button
import androidx.compose.material3.ButtonDefaults
import androidx.compose.material3.Card
import androidx.compose.material3.CardDefaults
import androidx.compose.material3.HorizontalDivider
import androidx.compose.material3.MaterialTheme
import androidx.compose.material3.OutlinedButton
import androidx.compose.material3.OutlinedTextField
import androidx.compose.material3.Surface
import androidx.compose.material3.Tab
import androidx.compose.material3.TabRow
import androidx.compose.material3.Text
import androidx.compose.material3.TextButton
import androidx.compose.runtime.Composable
import androidx.compose.runtime.getValue
import androidx.compose.runtime.mutableStateMapOf
import androidx.compose.runtime.mutableStateOf
import androidx.compose.runtime.remember
import androidx.compose.runtime.setValue
import androidx.compose.ui.Alignment
import androidx.compose.ui.Modifier
import androidx.compose.ui.graphics.Color
import androidx.compose.ui.text.TextStyle
import androidx.compose.ui.text.font.FontWeight
import androidx.compose.ui.text.input.ImeAction
import androidx.compose.ui.text.input.KeyboardType
import androidx.compose.ui.text.style.TextAlign
import androidx.compose.ui.unit.dp
import androidx.compose.ui.unit.sp
import com.example.app.excel.MaterialCatalog
import com.example.app.excel.MaterialCatalogReader
import com.example.app.excel.MaterialUnit
import org.apache.poi.ss.usermodel.BorderStyle
import org.apache.poi.ss.usermodel.FillPatternType
import org.apache.poi.ss.usermodel.HorizontalAlignment
import org.apache.poi.ss.usermodel.IndexedColors
import org.apache.poi.ss.usermodel.VerticalAlignment
import org.apache.poi.ss.util.CellRangeAddress
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import org.apache.poi.xssf.usermodel.XSSFCellStyle
import java.text.SimpleDateFormat
import java.util.Date
import java.util.Locale

private val Teal = Color(0xFF1699A8)
private val Navy = Color(0xFF0B6672)
private val Page = Color(0xFFF7F3EA)
private val RowAlt = Color(0xFFE4F5F6)
private val Green = Color(0xFF197D83)

private data class ProductLine(val code: String, val name: String, val carton: MaterialCatalog?, val box: MaterialCatalog?)
private data class Totals(val wholesale: Double, val retail: Double, val revenue: Double)
private enum class SalesTab(val title: String) { WHOLESALE("مبيع جملة"), RETAIL("مبيع مفرق"), TOTAL("الكلي") }

class MainActivity : ComponentActivity() {
    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)
        setContent { SalesMovementApp() }
    }
}

@Composable
private fun SalesMovementApp() {
    val context = androidx.compose.ui.platform.LocalContext.current
    val preferences = remember { context.getSharedPreferences("sales_movement_data", Context.MODE_PRIVATE) }
    val date = remember { SimpleDateFormat("dd/MM/yyyy", Locale.US).format(Date()) }
    val dateKey = remember { SimpleDateFormat("yyyyMMdd", Locale.US).format(Date()) }
    var catalog by remember { mutableStateOf(loadCatalog(preferences)) }
    val quantities = remember { mutableStateMapOf<String, String>().apply { putAll(loadDay(preferences, dateKey)) } }
    var message by remember { mutableStateOf("") }
    var historyOpen by remember { mutableStateOf(false) }
    var selectedTab by remember { mutableStateOf(SalesTab.TOTAL) }
    val launcher = rememberLauncherForActivityResult(ActivityResultContracts.OpenDocument()) { uri ->
        if (uri == null) return@rememberLauncherForActivityResult
        runCatching {
            val items = MaterialCatalogReader(context.contentResolver).read(uri)
            saveCatalog(preferences, items)
            catalog = items
            message = "تم تحميل ${items.map { it.code }.distinct().size} مادة"
        }.onFailure { message = it.message ?: "فشل قراءة دليل المواد" }
    }
    val exporter = rememberLauncherForActivityResult(ActivityResultContracts.CreateDocument("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")) { uri ->
        if (uri != null) runCatching { exportXlsx(context, uri, date, catalog, quantities) }.onSuccess { message = "تم تصدير Excel" }.onFailure { message = "فشل التصدير: ${it.message}" }
    }
    val historyExporter = rememberLauncherForActivityResult(ActivityResultContracts.CreateDocument("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")) { uri ->
        if (uri != null) runCatching { exportHistoryXlsx(context, uri, preferences) }.onSuccess { message = "تم تصدير سجل المبيعات" }.onFailure { message = "فشل تصدير السجل: ${it.message}" }
    }
    val products = catalog.groupBy { it.code }.map { (code, items) -> ProductLine(code, items.first().name, items.firstOrNull { it.unit == MaterialUnit.CARTON }, items.firstOrNull { it.unit == MaterialUnit.BOX }) }.sortedBy { it.code }
    val totals = calculateTotals(products, quantities)

    MaterialTheme {
        Surface(Modifier.fillMaxSize(), color = Page) {
            Column(Modifier.fillMaxSize()) {
                Header(date)
                SummaryBar(totals)
                Actions(
                    onImport = { launcher.launch(arrayOf("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "application/vnd.ms-excel")) },
                    onSave = { saveDay(preferences, dateKey, quantities, totals); message = "تم حفظ مبيعات $date" },
                    onExport = { exporter.launch("مبيعات-${date.replace('/', '-')}.xlsx") },
                    onExportHistory = { historyExporter.launch("سجل-المبيعات.xlsx") },
                    onHistory = { historyOpen = true },
                    onClear = { quantities.clear(); message = "تم تصفير إدخالات اليوم" },
                )
                if (message.isNotBlank()) Text(message, color = Green, fontSize = 14.sp, modifier = Modifier.padding(horizontal = 16.dp, vertical = 4.dp))
                if (products.isEmpty()) EmptyState() else {
                    TabRow(selectedTabIndex = selectedTab.ordinal, containerColor = Color.White, contentColor = Teal) {
                        SalesTab.entries.forEach { tab ->
                            Tab(selected = selectedTab == tab, onClick = { selectedTab = tab }, modifier = Modifier.height(36.dp), text = { Text(tab.title, fontSize = 13.sp, fontWeight = FontWeight.Bold, maxLines = 1) })
                        }
                    }
                    TableHeader(selectedTab)
                    LazyColumn(Modifier.fillMaxSize().padding(horizontal = 10.dp)) {
                        items(products) { product -> ProductRow(product, selectedTab, quantities) { field, value -> quantities[key(product.code, field)] = value.filter { it.isDigit() } } }
                        item { Spacer(Modifier.height(28.dp)) }
                    }
                }
            }
        }
    }
    if (historyOpen) HistoryDialog(preferences) { historyOpen = false }
}

@Composable
private fun Header(date: String) {
    Column(Modifier.fillMaxWidth().background(Teal).padding(horizontal = 14.dp, vertical = 9.dp), horizontalAlignment = Alignment.End) {
        Text("مبيعاتي اليومية", color = Color.White, fontSize = 23.sp, fontWeight = FontWeight.Bold)
        Text("الجملة والمفرق — $date", color = Color.White, fontSize = 13.sp)
    }
}

@Composable
private fun SummaryBar(t: Totals) {
    Card(Modifier.fillMaxWidth().padding(8.dp), colors = CardDefaults.cardColors(containerColor = Color.White), shape = RoundedCornerShape(14.dp)) {
        Column(Modifier.padding(horizontal = 8.dp, vertical = 7.dp)) {
            Row(Modifier.fillMaxWidth(), horizontalArrangement = Arrangement.SpaceEvenly) {
                SummaryCell("جملة", t.wholesale)
                SummaryCell("مفرق", t.retail)
                SummaryCell("الإجمالي", t.revenue, Green)
            }
            HorizontalDivider(Modifier.padding(vertical = 8.dp), color = Color(0xFFD7E2E6))
        }
    }
}

@Composable
private fun SummaryCell(label: String, value: Double, color: Color = Navy, suffix: String = "") {
    Column(horizontalAlignment = Alignment.CenterHorizontally) {
        Text(label, color = Color(0xFF687782), fontSize = 11.sp)
        Text(formatMoney(value) + suffix, color = color, fontSize = 15.sp, fontWeight = FontWeight.Bold)
    }
}

@Composable
private fun Actions(onImport: () -> Unit, onSave: () -> Unit, onExport: () -> Unit, onExportHistory: () -> Unit, onHistory: () -> Unit, onClear: () -> Unit) {
    Row(Modifier.fillMaxWidth().padding(horizontal = 8.dp), horizontalArrangement = Arrangement.spacedBy(3.dp)) {
        Button(onClick = onImport, colors = ButtonDefaults.buttonColors(containerColor = Teal), contentPadding = ButtonDefaults.TextButtonContentPadding, modifier = Modifier.weight(1f).height(40.dp)) { Text("دليل", fontSize = 12.sp, maxLines = 1) }
        Button(onClick = onSave, colors = ButtonDefaults.buttonColors(containerColor = Green), contentPadding = ButtonDefaults.TextButtonContentPadding, modifier = Modifier.weight(1f).height(40.dp)) { Text("حفظ اليوم", fontSize = 11.sp, maxLines = 1) }
        OutlinedButton(onClick = onExport, contentPadding = ButtonDefaults.TextButtonContentPadding, modifier = Modifier.weight(1f).height(40.dp)) { Text("Excel", fontSize = 12.sp, maxLines = 1) }
        TextButton(onClick = onHistory, contentPadding = ButtonDefaults.TextButtonContentPadding, modifier = Modifier.weight(1f).height(40.dp)) { Text("السجل", fontSize = 12.sp, maxLines = 1) }
    }
    OutlinedButton(onClick = onExportHistory, contentPadding = ButtonDefaults.TextButtonContentPadding, modifier = Modifier.fillMaxWidth().padding(horizontal = 8.dp).height(34.dp)) { Text("تصدير سجل Excel", fontSize = 12.sp, maxLines = 1) }
    TextButton(onClick = onClear, modifier = Modifier.fillMaxWidth().height(32.dp)) { Text("تصفير الإدخالات", fontSize = 12.sp) }
}

@Composable
private fun EmptyState() {
    Card(Modifier.padding(16.dp).fillMaxWidth(), colors = CardDefaults.cardColors(containerColor = Color.White), shape = RoundedCornerShape(16.dp)) {
        Column(Modifier.padding(24.dp), horizontalAlignment = Alignment.End) {
            Text("لم يتم تحميل دليل المواد بعد", color = Navy, fontSize = 20.sp, fontWeight = FontWeight.Bold)
            Spacer(Modifier.height(8.dp))
            Text("اضغط «دليل» وارفع ملف Excel المختصر.", color = Navy, fontSize = 15.sp)
        }
    }
}

@Composable
private fun TableHeader(tab: SalesTab) {
    Row(Modifier.fillMaxWidth().background(Navy).padding(vertical = 4.dp, horizontal = 3.dp), verticalAlignment = Alignment.CenterVertically) {
        Text("المادة", color = Color.White, fontSize = 12.sp, fontWeight = FontWeight.Bold, modifier = Modifier.weight(1.55f), textAlign = TextAlign.Center, maxLines = 1)
        if (tab != SalesTab.RETAIL) Text("طرد", color = Color.White, fontSize = 12.sp, fontWeight = FontWeight.Bold, modifier = Modifier.weight(1f), textAlign = TextAlign.Center, maxLines = 1)
        if (tab != SalesTab.RETAIL) Text("علبة", color = Color.White, fontSize = 12.sp, fontWeight = FontWeight.Bold, modifier = Modifier.weight(1f), textAlign = TextAlign.Center, maxLines = 1)
        if (tab != SalesTab.WHOLESALE) Text("طرد", color = Color.White, fontSize = 12.sp, fontWeight = FontWeight.Bold, modifier = Modifier.weight(1f), textAlign = TextAlign.Center, maxLines = 1)
        if (tab != SalesTab.WHOLESALE) Text("علبة", color = Color.White, fontSize = 12.sp, fontWeight = FontWeight.Bold, modifier = Modifier.weight(1f), textAlign = TextAlign.Center, maxLines = 1)
    }
}

@Composable
private fun ProductRow(product: ProductLine, tab: SalesTab, quantities: MutableMap<String, String>, onChange: (String, String) -> Unit) {
    val rowColor = if (product.code.hashCode() % 2 == 0) Color.White else RowAlt
    Row(Modifier.fillMaxWidth().background(rowColor).padding(vertical = 5.dp, horizontal = 3.dp), verticalAlignment = Alignment.CenterVertically) {
        Column(Modifier.weight(1.55f).padding(start = 7.dp, end = 3.dp), horizontalAlignment = Alignment.End) {
            Box(Modifier.fillMaxWidth(), contentAlignment = Alignment.CenterEnd) {
                Text(product.name, color = Navy, fontSize = 13.sp, fontWeight = FontWeight.Bold, maxLines = 1, modifier = Modifier.align(Alignment.CenterEnd))
                Text(priceLabel(product, tab), color = Color(0xFF687782), fontSize = 8.sp, maxLines = 1, modifier = Modifier.align(Alignment.CenterStart))
            }
        }
        if (tab != SalesTab.RETAIL) QuantityCell(product.carton != null, quantities[key(product.code, "carton_wholesale")].orEmpty(), Modifier.weight(1f)) { onChange("carton_wholesale", it) }
        if (tab != SalesTab.RETAIL) QuantityCell(product.box != null, quantities[key(product.code, "box_wholesale")].orEmpty(), Modifier.weight(1f)) { onChange("box_wholesale", it) }
        if (tab != SalesTab.WHOLESALE) QuantityCell(product.carton != null, quantities[key(product.code, "carton_retail")].orEmpty(), Modifier.weight(1f)) { onChange("carton_retail", it) }
        if (tab != SalesTab.WHOLESALE) QuantityCell(product.box != null, quantities[key(product.code, "box_retail")].orEmpty(), Modifier.weight(1f)) { onChange("box_retail", it) }
    }
    HorizontalDivider(color = Color(0xFFD7E2E6))
}

@Composable
private fun QuantityCell(enabled: Boolean, value: String, modifier: Modifier, onValueChange: (String) -> Unit) {
    OutlinedTextField(value = value, onValueChange = onValueChange, enabled = enabled, singleLine = true, keyboardOptions = androidx.compose.foundation.text.KeyboardOptions(keyboardType = KeyboardType.Number, imeAction = ImeAction.Done), textStyle = TextStyle(fontSize = 13.sp, textAlign = TextAlign.Center), modifier = modifier.height(42.dp).padding(horizontal = 2.dp), placeholder = { Text("0", fontSize = 12.sp) })
}

private fun priceLabel(product: ProductLine, tab: SalesTab): String {
    fun price(value: Double?): String = value?.let { formatMoney(it) } ?: "—"
    return when (tab) {
        SalesTab.WHOLESALE -> "${price(product.carton?.wholesalePrice)}/${price(product.box?.wholesalePrice)}"
        SalesTab.RETAIL -> "${price(product.carton?.retailPrice)}/${price(product.box?.retailPrice)}"
        SalesTab.TOTAL -> "ج${price(product.carton?.wholesalePrice)} م${price(product.carton?.retailPrice)}"
    }
}

@Composable
private fun HistoryDialog(preferences: android.content.SharedPreferences, close: () -> Unit) {
    val days = preferences.getStringSet("saved_days", emptySet()).orEmpty().toList().sortedDescending()
    AlertDialog(onDismissRequest = close, confirmButton = { TextButton(onClick = close) { Text("إغلاق") } }, title = { Text("سجل الأيام", color = Navy, fontWeight = FontWeight.Bold) }, text = {
        if (days.isEmpty()) Text("لا يوجد أيام محفوظة بعد.") else LazyColumn { items(days) { day ->
            val parts = preferences.getString("summary_$day", "0|0|0").orEmpty().split('|')
            Row(Modifier.fillMaxWidth().padding(vertical = 6.dp), horizontalArrangement = Arrangement.SpaceBetween) {
                Text(dateLabel(day), color = Navy, fontWeight = FontWeight.Bold)
                Text("إجمالي المبيعات ${parts.getOrNull(0) ?: "0"}", color = Green)
            }
            HorizontalDivider()
        } }
    })
}

private fun key(code: String, field: String) = "$code:$field"
private fun amount(value: String?): Double = value?.toDoubleOrNull() ?: 0.0
private fun formatMoney(value: Double): String = String.format(Locale.US, "%,.0f", value)
private fun dateLabel(key: String): String = runCatching { SimpleDateFormat("dd/MM/yyyy", Locale.US).format(SimpleDateFormat("yyyyMMdd", Locale.US).parse(key)!!) }.getOrDefault(key)

private fun calculateTotals(products: List<ProductLine>, q: Map<String, String>): Totals {
    var wholesale = 0.0; var retail = 0.0
    products.forEach { p ->
        val entries: List<Pair<Double, MaterialCatalog?>> = listOf(
            Pair(amount(q[key(p.code, "carton_wholesale")]), p.carton),
            Pair(amount(q[key(p.code, "box_wholesale")]), p.box),
            Pair(amount(q[key(p.code, "carton_retail")]), p.carton),
            Pair(amount(q[key(p.code, "box_retail")]), p.box)
        )
        entries.forEachIndexed { index, (qty, item) ->
            val price = if (index < 2) item?.wholesalePrice else item?.retailPrice
            val value = kotlin.math.round(qty * (price ?: 0.0))
            if (index < 2) wholesale += value else retail += value
        }
    }
    return Totals(wholesale, retail, wholesale + retail)
}

private fun saveCatalog(p: android.content.SharedPreferences, items: List<MaterialCatalog>) = p.edit().putString("catalog", items.joinToString("\n") { listOf(it.code, it.name, it.unit.label, it.costPrice ?: "", it.wholesalePrice ?: "", it.retailPrice ?: "").joinToString("|") }).apply()
private fun loadCatalog(p: android.content.SharedPreferences): List<MaterialCatalog> = p.getString("catalog", "").orEmpty().lineSequence().mapNotNull { line ->
    val a = line.split('|'); if (a.size < 6) return@mapNotNull null
    val unit = MaterialUnit.from(a[2]); if (unit == MaterialUnit.OTHER) return@mapNotNull null
    MaterialCatalog(a[0], a[1], unit, a[3].toDoubleOrNull(), a[4].toDoubleOrNull(), a[5].toDoubleOrNull())
}.toList()
private fun saveDay(p: android.content.SharedPreferences, day: String, q: Map<String, String>, t: Totals) { p.edit().putString("day_$day", q.entries.joinToString("\n") { "${it.key}=${it.value}" }).putString("summary_$day", "${t.revenue}").putStringSet("saved_days", p.getStringSet("saved_days", emptySet()).orEmpty() + day).apply() }
private fun loadDay(p: android.content.SharedPreferences, day: String): Map<String, String> = p.getString("day_$day", "").orEmpty().lineSequence().mapNotNull { line -> line.split('=', limit = 2).takeIf { it.size == 2 }?.let { it[0] to it[1] } }.toMap()

private fun exportXlsx(context: Context, uri: android.net.Uri, date: String, catalog: List<MaterialCatalog>, q: Map<String, String>) {
    val products = catalog.groupBy { it.code }.map { (code, items) -> ProductLine(code, items.first().name, items.firstOrNull { it.unit == MaterialUnit.CARTON }, items.firstOrNull { it.unit == MaterialUnit.BOX }) }.sortedBy { it.code }
    val workbook = XSSFWorkbook()
    val sheet = workbook.createSheet("مبيعات ${date.replace('/', '-')}")
    sheet.setRightToLeft(true)
    val headers = listOf("رمز المادة", "المادة", "جملة طرد", "جملة علبة", "مفرق طرد", "مفرق علبة", "قيمة الجملة", "قيمة المفرق", "إجمالي المبيعات")
    val numberFormat = workbook.createDataFormat().getFormat("#,##0")
    fun makeStyle(fill: IndexedColors, bold: Boolean = false, numeric: Boolean = false): XSSFCellStyle = workbook.createCellStyle().apply {
        fillForegroundColor = fill.index
        fillPattern = FillPatternType.SOLID_FOREGROUND
        alignment = HorizontalAlignment.CENTER
        verticalAlignment = VerticalAlignment.CENTER
        borderLeft = BorderStyle.THIN
        borderRight = BorderStyle.THIN
        borderTop = BorderStyle.THIN
        borderBottom = BorderStyle.THIN
        if (numeric) dataFormat = numberFormat
        val font = workbook.createFont().apply { this.bold = bold; fontHeightInPoints = if (bold) 11 else 10; color = if (fill == IndexedColors.DARK_TEAL) IndexedColors.WHITE.index else IndexedColors.BLACK.index }
        setFont(font)
    }
    val headerStyle = makeStyle(IndexedColors.DARK_TEAL, bold = true)
    val bodyStyle = makeStyle(IndexedColors.WHITE)
    val altStyle = makeStyle(IndexedColors.LIGHT_TURQUOISE)
    val bodyNumberStyle = makeStyle(IndexedColors.WHITE, numeric = true)
    val altNumberStyle = makeStyle(IndexedColors.LIGHT_TURQUOISE, numeric = true)
    val totalStyle = makeStyle(IndexedColors.LIGHT_TURQUOISE, bold = true, numeric = true)
    val header = sheet.createRow(0)
    header.heightInPoints = 26f
    headers.forEachIndexed { i, h -> header.createCell(i).apply { setCellValue(h); cellStyle = headerStyle } }
    var excelRow = 1
    var wholesaleTotal = 0.0
    var retailTotal = 0.0
    products.forEach { p ->
        val quantities = listOf("carton_wholesale", "box_wholesale", "carton_retail", "box_retail").map { amount(q[key(p.code, it)]) }
        if (quantities.all { it == 0.0 }) return@forEach
        val row = sheet.createRow(excelRow++)
        row.createCell(0).setCellValue(p.code)
        row.createCell(1).setCellValue(p.name)
        val wholesaleValue = kotlin.math.round(quantities[0] * (p.carton?.wholesalePrice ?: 0.0) + quantities[1] * (p.box?.wholesalePrice ?: 0.0))
        val retailValue = kotlin.math.round(quantities[2] * (p.carton?.retailPrice ?: 0.0) + quantities[3] * (p.box?.retailPrice ?: 0.0))
        wholesaleTotal += wholesaleValue
        retailTotal += retailValue
        val textStyle = if (excelRow % 2 == 0) altStyle else bodyStyle
        val numericStyle = if (excelRow % 2 == 0) altNumberStyle else bodyNumberStyle
        row.getCell(0).cellStyle = textStyle
        row.createCell(1).apply { setCellValue(p.name); cellStyle = textStyle }
        quantities.forEachIndexed { i, value -> row.createCell(i + 2).apply { setCellValue(value); cellStyle = numericStyle } }
        row.createCell(6).apply { setCellValue(wholesaleValue); cellStyle = numericStyle }
        row.createCell(7).apply { setCellValue(retailValue); cellStyle = numericStyle }
        row.createCell(8).apply { setCellValue(wholesaleValue + retailValue); cellStyle = numericStyle }
    }
    val total = sheet.createRow(excelRow)
    total.createCell(0).apply { setCellValue("الإجمالي"); cellStyle = totalStyle }
    total.createCell(1).cellStyle = totalStyle
    for (i in 2..5) total.createCell(i).apply { setCellValue(quantitiesTotal(products, q, i - 2)); cellStyle = totalStyle }
    total.createCell(6).apply { setCellValue(wholesaleTotal); cellStyle = totalStyle }
    total.createCell(7).apply { setCellValue(retailTotal); cellStyle = totalStyle }
    total.createCell(8).apply { setCellValue(wholesaleTotal + retailTotal); cellStyle = totalStyle }
    sheet.createFreezePane(0, 1)
    sheet.setAutoFilter(CellRangeAddress(0, excelRow, 0, 8))
    listOf(14, 26, 12, 12, 12, 12, 16, 16, 18).forEachIndexed { i, width -> sheet.setColumnWidth(i, width * 256) }
    context.contentResolver.openOutputStream(uri)?.use { workbook.write(it) } ?: error("تعذر إنشاء ملف Excel")
    workbook.close()
}

private fun quantitiesTotal(products: List<ProductLine>, q: Map<String, String>, index: Int): Double = products.sumOf { p -> amount(q[key(p.code, listOf("carton_wholesale", "box_wholesale", "carton_retail", "box_retail")[index])]) }

private fun exportHistoryXlsx(context: Context, uri: android.net.Uri, preferences: android.content.SharedPreferences) {
    val workbook = XSSFWorkbook()
    val sheet = workbook.createSheet("سجل المبيعات")
    sheet.setRightToLeft(true)
    sheet.createRow(0).apply {
        createCell(0).setCellValue("التاريخ")
        createCell(1).setCellValue("إجمالي المبيعات")
    }
    preferences.getStringSet("saved_days", emptySet()).orEmpty().toList().sorted().forEachIndexed { index, day ->
        val row = sheet.createRow(index + 1)
        row.createCell(0).setCellValue(dateLabel(day))
        row.createCell(1).setCellValue(preferences.getString("summary_$day", "0").orEmpty().split('|').firstOrNull()?.toDoubleOrNull() ?: 0.0)
    }
    context.contentResolver.openOutputStream(uri)?.use { workbook.write(it) } ?: error("تعذر إنشاء ملف سجل Excel")
    workbook.close()
}
