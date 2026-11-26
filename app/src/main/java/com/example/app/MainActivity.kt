package com.example.app

import android.content.Context
import android.os.Bundle
import androidx.activity.ComponentActivity
import androidx.activity.compose.rememberLauncherForActivityResult
import androidx.activity.compose.setContent
import androidx.activity.result.contract.ActivityResultContracts
import androidx.compose.foundation.background
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
import androidx.compose.ui.text.style.TextAlign
import androidx.compose.ui.unit.dp
import androidx.compose.ui.unit.sp
import com.example.app.excel.MaterialCatalog
import com.example.app.excel.MaterialCatalogReader
import com.example.app.excel.MaterialUnit
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.text.SimpleDateFormat
import java.util.Date
import java.util.Locale

private val Teal = Color(0xFF1196A7)
private val Navy = Color(0xFF10283D)
private val Page = Color(0xFFF7F3EA)
private val RowAlt = Color(0xFFF4F7F8)
private val Green = Color(0xFF2E7D50)

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
                    onExport = { exporter.launch("مبيعات-$date.xlsx") },
                    onHistory = { historyOpen = true },
                    onClear = { quantities.clear(); message = "تم تصفير إدخالات اليوم" },
                )
                if (message.isNotBlank()) Text(message, color = Green, fontSize = 14.sp, modifier = Modifier.padding(horizontal = 16.dp, vertical = 4.dp))
                if (products.isEmpty()) EmptyState() else {
                    TabRow(selectedTabIndex = selectedTab.ordinal, containerColor = Color.White, contentColor = Teal) {
                        SalesTab.entries.forEach { tab ->
                            Tab(selected = selectedTab == tab, onClick = { selectedTab = tab }, text = { Text(tab.title, fontWeight = FontWeight.Bold) })
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
    Column(Modifier.fillMaxWidth().background(Teal).padding(18.dp), horizontalAlignment = Alignment.End) {
        Text("مبيعاتي اليومية", color = Color.White, fontSize = 27.sp, fontWeight = FontWeight.Bold)
        Text("الجملة والمفرق — $date", color = Color.White, fontSize = 15.sp)
    }
}

@Composable
private fun SummaryBar(t: Totals) {
    Card(Modifier.fillMaxWidth().padding(12.dp), colors = CardDefaults.cardColors(containerColor = Color.White), shape = RoundedCornerShape(14.dp)) {
        Column(Modifier.padding(12.dp)) {
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
        Text(label, color = Color(0xFF687782), fontSize = 12.sp)
        Text(formatMoney(value) + suffix, color = color, fontSize = 16.sp, fontWeight = FontWeight.Bold)
    }
}

@Composable
private fun Actions(onImport: () -> Unit, onSave: () -> Unit, onExport: () -> Unit, onHistory: () -> Unit, onClear: () -> Unit) {
    Row(Modifier.fillMaxWidth().padding(horizontal = 12.dp), horizontalArrangement = Arrangement.spacedBy(6.dp)) {
        Button(onClick = onImport, colors = ButtonDefaults.buttonColors(containerColor = Teal), modifier = Modifier.weight(1f)) { Text("دليل") }
        Button(onClick = onSave, colors = ButtonDefaults.buttonColors(containerColor = Green), modifier = Modifier.weight(1f)) { Text("حفظ اليوم") }
        OutlinedButton(onClick = onExport, modifier = Modifier.weight(1f)) { Text("Excel") }
        TextButton(onClick = onHistory, modifier = Modifier.weight(1f)) { Text("السجل") }
    }
    TextButton(onClick = onClear, modifier = Modifier.fillMaxWidth()) { Text("تصفير الإدخالات") }
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
    Row(Modifier.fillMaxWidth().background(Navy).padding(vertical = 9.dp, horizontal = 4.dp), verticalAlignment = Alignment.CenterVertically) {
        Text("المادة", color = Color.White, fontWeight = FontWeight.Bold, modifier = Modifier.weight(1.55f), textAlign = TextAlign.Center)
        if (tab != SalesTab.RETAIL) Text("جملة\nطرد", color = Color.White, fontWeight = FontWeight.Bold, modifier = Modifier.weight(1f), textAlign = TextAlign.Center)
        if (tab != SalesTab.RETAIL) Text("جملة\nعلبة", color = Color.White, fontWeight = FontWeight.Bold, modifier = Modifier.weight(1f), textAlign = TextAlign.Center)
        if (tab != SalesTab.WHOLESALE) Text("مفرق\nطرد", color = Color.White, fontWeight = FontWeight.Bold, modifier = Modifier.weight(1f), textAlign = TextAlign.Center)
        if (tab != SalesTab.WHOLESALE) Text("مفرق\nعلبة", color = Color.White, fontWeight = FontWeight.Bold, modifier = Modifier.weight(1f), textAlign = TextAlign.Center)
    }
}

@Composable
private fun ProductRow(product: ProductLine, tab: SalesTab, quantities: MutableMap<String, String>, onChange: (String, String) -> Unit) {
    val rowColor = if (product.code.hashCode() % 2 == 0) Color.White else RowAlt
    Row(Modifier.fillMaxWidth().background(rowColor).padding(vertical = 7.dp, horizontal = 3.dp), verticalAlignment = Alignment.CenterVertically) {
        Column(Modifier.weight(1.55f).padding(horizontal = 3.dp), horizontalAlignment = Alignment.End) {
            Text(product.name, color = Navy, fontSize = 13.sp, fontWeight = FontWeight.Bold, maxLines = 1)
            Text(product.code, color = Color(0xFF687782), fontSize = 11.sp)
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
    OutlinedTextField(value = value, onValueChange = onValueChange, enabled = enabled, singleLine = true, textStyle = TextStyle(fontSize = 13.sp, textAlign = TextAlign.Center), modifier = modifier.padding(horizontal = 2.dp), placeholder = { Text("0", fontSize = 12.sp) })
}

@Composable
private fun HistoryDialog(preferences: android.content.SharedPreferences, close: () -> Unit) {
    val days = preferences.getStringSet("saved_days", emptySet()).orEmpty().toList().sortedDescending()
    AlertDialog(onDismissRequest = close, confirmButton = { TextButton(onClick = close) { Text("إغلاق") } }, title = { Text("سجل الأيام", color = Navy, fontWeight = FontWeight.Bold) }, text = {
        if (days.isEmpty()) Text("لا يوجد أيام محفوظة بعد.") else LazyColumn { items(days) { day ->
            val parts = preferences.getString("summary_$day", "0|0|0").orEmpty().split('|')
            Row(Modifier.fillMaxWidth().padding(vertical = 6.dp), horizontalArrangement = Arrangement.SpaceBetween) {
                Text(day, color = Navy, fontWeight = FontWeight.Bold)
                Text("إجمالي المبيعات ${parts.getOrNull(0) ?: "0"}", color = Green)
            }
            HorizontalDivider()
        } }
    })
}

private fun key(code: String, field: String) = "$code:$field"
private fun amount(value: String?): Double = value?.toDoubleOrNull() ?: 0.0
private fun formatMoney(value: Double): String = String.format(Locale.US, "%,.0f", value)

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
    val sheet = workbook.createSheet("مبيعات $date")
    sheet.setRightToLeft(true)
    val headers = listOf("رمز المادة", "المادة", "جملة طرد", "جملة علبة", "مفرق طرد", "مفرق علبة", "قيمة الجملة", "قيمة المفرق", "إجمالي المبيعات")
    headers.forEachIndexed { i, h -> sheet.createRow(0).createCell(i).setCellValue(h) }
    products.forEachIndexed { index, p ->
        val row = sheet.createRow(index + 1)
        row.createCell(0).setCellValue(p.code)
        row.createCell(1).setCellValue(p.name)
        val quantities = listOf("carton_wholesale", "box_wholesale", "carton_retail", "box_retail").map { amount(q[key(p.code, it)]) }
        quantities.forEachIndexed { i, value -> row.createCell(i + 2).setCellValue(value) }
        val wholesaleValue = kotlin.math.round(quantities[0] * (p.carton?.wholesalePrice ?: 0.0) + quantities[1] * (p.box?.wholesalePrice ?: 0.0))
        val retailValue = kotlin.math.round(quantities[2] * (p.carton?.retailPrice ?: 0.0) + quantities[3] * (p.box?.retailPrice ?: 0.0))
        row.createCell(6).setCellValue(wholesaleValue)
        row.createCell(7).setCellValue(retailValue)
        row.createCell(8).setCellValue(wholesaleValue + retailValue)
    }
    context.contentResolver.openOutputStream(uri)?.use { workbook.write(it) } ?: error("تعذر إنشاء ملف Excel")
    workbook.close()
}
