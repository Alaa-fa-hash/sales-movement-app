package com.example.app.excel

import android.content.ContentResolver
import android.net.Uri
import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.ss.usermodel.DataFormatter
import org.apache.poi.ss.usermodel.Row
import org.apache.poi.ss.usermodel.WorkbookFactory
import java.io.InputStream
import java.util.Locale

/** Reads the first visible sheet of the six-column material catalog. */
class MaterialCatalogReader(private val contentResolver: ContentResolver) {
    fun read(uri: Uri): List<MaterialCatalog> {
        val input = contentResolver.openInputStream(uri)
            ?: throw IllegalArgumentException("تعذر فتح ملف دليل المواد")
        return input.use { parse(it) }
    }

    private fun parse(stream: InputStream): List<MaterialCatalog> {
        WorkbookFactory.create(stream).use { workbook ->
            val sheet = workbook.getSheetAt(0)
            val formatter = DataFormatter(Locale.US)
            val iterator = sheet.iterator()
            if (!iterator.hasNext()) return emptyList()
            val headerRow = iterator.next()
            val headers = headerRow.asSequence().associate { normalize(cellText(it, formatter)) to it.columnIndex }
            val required = listOf("رمز المادة", "اسم المادة", "الوحدة", "سعر الجملة", "سعر المفرق")
            val missing = required.filter { normalize(it) !in headers }
            if (missing.isNotEmpty()) throw IllegalArgumentException("أعمدة دليل المواد الناقصة: ${missing.joinToString("، ")}")

            fun value(row: Row, header: String): String = cellText(row.getCell(headers.getValue(normalize(header))), formatter)
            fun optionalValue(row: Row, header: String): String = headers[normalize(header)]?.let { cellText(row.getCell(it), formatter) }.orEmpty()
            return buildList {
                while (iterator.hasNext()) {
                    val row = iterator.next()
                    val code = value(row, "رمز المادة").trim()
                    val name = value(row, "اسم المادة").trim()
                    val unit = MaterialUnit.from(value(row, "الوحدة"))
                    if (code.isBlank() || name.isBlank() || unit == MaterialUnit.OTHER) continue
                    add(MaterialCatalog(code, name, unit, parsePrice(optionalValue(row, "سعر التكلفة")), parsePrice(value(row, "سعر الجملة")), parsePrice(value(row, "سعر المفرق"))))
                }
            }
        }
    }

    private fun cellText(cell: Cell?, formatter: DataFormatter): String =
        if (cell == null || cell.cellType == Cell.CELL_TYPE_BLANK) "" else formatter.formatCellValue(cell).trim()

    private fun parsePrice(value: String): Double? = value.replace(",", "").trim().toDoubleOrNull()

    private fun normalize(value: String): String = value.trim().replace(" ", "").replace("ـ", "").lowercase(Locale.ROOT)
}
