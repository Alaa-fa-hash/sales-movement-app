package com.example.app

import android.content.Context
import android.os.Bundle
import androidx.activity.ComponentActivity
import androidx.activity.compose.rememberLauncherForActivityResult
import androidx.activity.compose.setContent
import androidx.activity.result.contract.ActivityResultContracts
import androidx.compose.foundation.background
import androidx.compose.foundation.interaction.MutableInteractionSource
import androidx.compose.foundation.layout.Box
import androidx.compose.foundation.layout.Arrangement
import androidx.compose.foundation.layout.Column
import androidx.compose.foundation.layout.PaddingValues
import androidx.compose.foundation.layout.Row
import androidx.compose.foundation.layout.Spacer
import androidx.compose.foundation.layout.fillMaxSize
import androidx.compose.foundation.layout.fillMaxWidth
import androidx.compose.foundation.layout.height
import androidx.compose.foundation.layout.padding
import androidx.compose.foundation.layout.width
import androidx.compose.foundation.lazy.LazyColumn
import androidx.compose.foundation.lazy.items
import androidx.compose.foundation.lazy.itemsIndexed
import androidx.compose.foundation.text.BasicTextField
import androidx.compose.foundation.shape.RoundedCornerShape
import androidx.compose.material3.AlertDialog
import androidx.compose.material3.Button
import androidx.compose.material3.ButtonDefaults
import androidx.compose.material3.ExperimentalMaterial3Api
import androidx.compose.material3.Card
import androidx.compose.material3.CardDefaults
import androidx.compose.material3.HorizontalDivider
import androidx.compose.material3.MaterialTheme
import androidx.compose.material3.OutlinedButton
import androidx.compose.material3.OutlinedTextField
import androidx.compose.material3.OutlinedTextFieldDefaults
import androidx.compose.material3.Surface
import androidx.compose.material3.Tab
import androidx.compose.material3.TabRow
import androidx.compose.material3.Text
import androidx.compose.material3.TextButton
import androidx.compose.runtime.Composable
import androidx.compose.runtime.LaunchedEffect
import androidx.compose.runtime.getValue
import androidx.compose.runtime.mutableStateMapOf
import androidx.compose.runtime.mutableStateListOf
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
import androidx.compose.ui.text.input.VisualTransformation
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
import org.apache.poi.xssf.usermodel.XSSFColor
import java.io.ByteArrayOutputStream
import java.text.SimpleDateFormat
import java.util.Date
import java.util.Locale

private val Teal = Color(0xFF0B7A8C)
private val Navy = Color(0xFF0B7A8C)
private val Page = Color(0xFFFFFFFF)
private val RowAlt = Color(0xFFE8F4F6)
private val TotalFill = Color(0xFFD5ECEF)
private val Green = Color(0xFF0B7A8C)

private data class ProductLine(val code: String, val name: String, val carton: MaterialCatalog?, val box: MaterialCatalog?)
private data class Totals(val wholesale: Double, val retail: Double, val revenue: Double)
private class SpecialSaleDraft {
    var name by mutableStateOf("")
    var quantity by mutableStateOf("")
    var price by mutableStateOf("")
}
private enum class SalesTab(val title: String) { WHOLESALE("مبيع جملة"), RETAIL("مبيع مفرق"), TOTAL("الكلي"), SPECIAL("بيع خاص") }

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
    var newTaskConfirm by remember { mutableStateOf(false) }
    var selectedTab by remember { mutableStateOf(SalesTab.TOTAL) }
    val specialSales = remember {
        mutableStateListOf<SpecialSaleDraft>().apply {
            addAll(loadSpecialSales(preferences, dateKey))
            if (isEmpty()) add(SpecialSaleDraft())
        }
    }
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
        if (uri != null) runCatching { exportXlsx(context, uri, date, catalog, quantities, specialSales) }.onSuccess { message = "تم تصدير Excel" }.onFailure { message = "فشل التصدير: ${it.message}" }
    }
    val historyExporter = rememberLauncherForActivityResult(ActivityResultContracts.CreateDocument("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")) { uri ->
        if (uri != null) runCatching { exportHistoryXlsx(context, uri, preferences) }.onSuccess { message = "تم تصدير سجل المبيعات" }.onFailure { message = "فشل تصدير السجل: ${it.message}" }
    }
    val products = buildProducts(catalog)
    val totals = calculateTotals(products, quantities)
    val specialTotal = calculateSpecialTotal(specialSales)
    val dailyTotals = totals.copy(revenue = totals.revenue + specialTotal)
    fun persistCurrentDay() {
        val currentTotals = calculateTotals(products, quantities)
        saveDay(preferences, dateKey, quantities, currentTotals.copy(revenue = currentTotals.revenue + calculateSpecialTotal(specialSales)), specialSales)
    }
    LaunchedEffect(dateKey) {
        if (!preferences.getStringSet("saved_days", emptySet()).orEmpty().contains(dateKey)) {
            saveDay(preferences, dateKey, quantities, Totals(0.0, 0.0, 0.0), specialSales)
        }
    }

    MaterialTheme {
        Surface(Modifier.fillMaxSize(), color = Page) {
            Column(Modifier.fillMaxSize()) {
                Header(date)
                SummaryBar(dailyTotals)
                Actions(
                    onImport = { launcher.launch(arrayOf("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "application/vnd.ms-excel")) },
                    onSave = { persistCurrentDay(); message = "تم حفظ مبيعات $date" },
                        onExport = { exporter.launch("مبيعات-${date.replace('/', '-')}.xlsx") },
                    onExportHistory = { historyExporter.launch("سجل-المبيعات.xlsx") },
                    onHistory = { historyOpen = true },
                    onNewTask = { newTaskConfirm = true },
                    onClear = {
                        quantities.clear()
                        specialSales.clear()
                        specialSales.add(SpecialSaleDraft())
                        message = "تم تصفير الإدخالات دون تغيير السجل المحفوظ"
                    },
                )
                if (message.isNotBlank()) Text(message, color = Green, fontSize = 14.sp, modifier = Modifier.padding(horizontal = 16.dp, vertical = 4.dp))
                if (products.isEmpty()) EmptyState() else {
                    TabRow(selectedTabIndex = selectedTab.ordinal, containerColor = Color.White, contentColor = Teal) {
                        SalesTab.entries.forEach { tab ->
                            Tab(selected = selectedTab == tab, onClick = { selectedTab = tab }, modifier = Modifier.height(36.dp), text = { Text(tab.title, fontSize = 13.sp, fontWeight = FontWeight.Bold, maxLines = 1) })
                        }
                    }
                    if (selectedTab == SalesTab.SPECIAL) {
                    SpecialSalesTable(specialSales, onChanged = { persistCurrentDay() }) { specialSales.add(SpecialSaleDraft()); persistCurrentDay() }
                    } else {
                        TableHeader(selectedTab)
                        LazyColumn(Modifier.fillMaxSize().padding(horizontal = 10.dp)) {
                            itemsIndexed(products) { index, product -> ProductRow(product, selectedTab, quantities, index) { field, value -> quantities[key(product.code, field)] = sanitizeQuantity(value, field.startsWith("carton")); persistCurrentDay() } }
                            item { Spacer(Modifier.height(28.dp)) }
                        }
                    }
                }
            }
        }
    }
    if (historyOpen) HistoryDialog(preferences) { historyOpen = false }
    if (newTaskConfirm) {
        AlertDialog(
            onDismissRequest = { newTaskConfirm = false },
            title = { Text("بدء مهمة جديدة", color = Navy, fontWeight = FontWeight.Bold) },
            text = { Text("سيتم حذف كل سجلات المبيعات والبيع الخاص مع إبقاء دليل المواد. هل تريد المتابعة؟") },
            confirmButton = {
                Button(onClick = {
                    clearSalesHistory(preferences)
                    quantities.clear()
                    specialSales.clear()
                    specialSales.add(SpecialSaleDraft())
                    newTaskConfirm = false
                    message = "بدأت مهمة جديدة"
                }, colors = ButtonDefaults.buttonColors(containerColor = Teal)) { Text("بدء") }
            },
            dismissButton = { TextButton(onClick = { newTaskConfirm = false }) { Text("إلغاء") } },
        )
    }
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
    Column(horizontalAlignment = Alignment.CenterHorizontally, modifier = if (label == "الإجمالي") Modifier.background(TotalFill) else Modifier) {
        Text(label, color = Color(0xFF687782), fontSize = 11.sp)
        Text(formatMoney(value) + suffix, color = color, fontSize = 15.sp, fontWeight = FontWeight.Bold)
    }
}

@Composable
private fun Actions(onImport: () -> Unit, onSave: () -> Unit, onExport: () -> Unit, onExportHistory: () -> Unit, onHistory: () -> Unit, onNewTask: () -> Unit, onClear: () -> Unit) {
    Row(Modifier.fillMaxWidth().padding(horizontal = 8.dp), horizontalArrangement = Arrangement.spacedBy(3.dp)) {
        Button(onClick = onImport, colors = ButtonDefaults.buttonColors(containerColor = Teal), contentPadding = ButtonDefaults.TextButtonContentPadding, modifier = Modifier.weight(1f).height(40.dp)) { Text("دليل", fontSize = 12.sp, maxLines = 1) }
        Button(onClick = onSave, colors = ButtonDefaults.buttonColors(containerColor = Green), contentPadding = ButtonDefaults.TextButtonContentPadding, modifier = Modifier.weight(1f).height(40.dp)) { Text("حفظ اليوم", fontSize = 11.sp, maxLines = 1) }
        OutlinedButton(onClick = onExport, contentPadding = ButtonDefaults.TextButtonContentPadding, modifier = Modifier.weight(1f).height(40.dp)) { Text("Excel", fontSize = 12.sp, maxLines = 1) }
        TextButton(onClick = onHistory, contentPadding = ButtonDefaults.TextButtonContentPadding, modifier = Modifier.weight(1f).height(40.dp)) { Text("السجل", fontSize = 12.sp, maxLines = 1) }
    }
    OutlinedButton(onClick = onExportHistory, contentPadding = ButtonDefaults.TextButtonContentPadding, modifier = Modifier.fillMaxWidth().padding(horizontal = 8.dp).height(34.dp)) { Text("تصدير سجل Excel", fontSize = 12.sp, maxLines = 1) }
    TextButton(onClick = onClear, modifier = Modifier.fillMaxWidth().height(32.dp)) { Text("تصفير الإدخالات", fontSize = 12.sp) }
    Spacer(Modifier.height(22.dp))
    OutlinedButton(onClick = onNewTask, colors = ButtonDefaults.outlinedButtonColors(contentColor = Color(0xFFB3261E)), contentPadding = ButtonDefaults.TextButtonContentPadding, modifier = Modifier.fillMaxWidth().padding(horizontal = 8.dp).height(38.dp)) { Text("بدء مهمة جديدة", fontSize = 12.sp, maxLines = 1) }
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
@OptIn(ExperimentalMaterial3Api::class)
private fun SpecialSalesTable(rows: List<SpecialSaleDraft>, onChanged: () -> Unit, onAddRow: () -> Unit) {
    Column(Modifier.fillMaxSize().padding(horizontal = 10.dp)) {
        Row(Modifier.fillMaxWidth().background(Navy).padding(vertical = 4.dp, horizontal = 3.dp), verticalAlignment = Alignment.CenterVertically) {
            Text("المادة", color = Color.White, fontSize = 12.sp, fontWeight = FontWeight.Bold, modifier = Modifier.weight(1.8f), textAlign = TextAlign.Center, maxLines = 1)
            Text("الكمية", color = Color.White, fontSize = 12.sp, fontWeight = FontWeight.Bold, modifier = Modifier.weight(1f), textAlign = TextAlign.Center, maxLines = 1)
            Text("السعر الخاص", color = Color.White, fontSize = 11.sp, fontWeight = FontWeight.Bold, modifier = Modifier.weight(1.15f), textAlign = TextAlign.Center, maxLines = 1)
            Text("الإجمالي", color = Color.White, fontSize = 11.sp, fontWeight = FontWeight.Bold, modifier = Modifier.weight(1.15f), textAlign = TextAlign.Center, maxLines = 1)
        }
        LazyColumn(Modifier.fillMaxSize()) {
            itemsIndexed(rows) { index, row ->
                val total = if (row.quantity.isNotBlank() && row.price.isNotBlank()) amount(row.quantity) * amount(row.price) else null
                val rowColor = if (index % 2 == 0) Color.White else RowAlt
                Row(Modifier.fillMaxWidth().background(rowColor).padding(vertical = 5.dp, horizontal = 2.dp), verticalAlignment = Alignment.CenterVertically) {
                    SpecialInputCell(row.name, Modifier.weight(1.8f), KeyboardType.Text, "اسم المادة") { row.name = it; onChanged() }
                    SpecialInputCell(row.quantity, Modifier.weight(1f), KeyboardType.Decimal, "كمية") { row.quantity = sanitizeQuantity(it, true); onChanged() }
                    SpecialInputCell(row.price, Modifier.weight(1.15f), KeyboardType.Decimal, "سعر") { row.price = sanitizeQuantity(it, true); onChanged() }
                    Text(total?.let(::formatMoney) ?: "—", color = Navy, fontSize = 12.sp, fontWeight = FontWeight.Bold, maxLines = 1, modifier = Modifier.weight(1.15f), textAlign = TextAlign.Center)
                }
                HorizontalDivider(color = Color(0xFFD7E2E6))
            }
            item {
                OutlinedButton(onClick = onAddRow, modifier = Modifier.fillMaxWidth().padding(vertical = 10.dp)) { Text("+ إضافة سطر") }
                Spacer(Modifier.height(28.dp))
            }
        }
    }
}

@Composable
@OptIn(ExperimentalMaterial3Api::class)
private fun SpecialInputCell(value: String, modifier: Modifier, keyboardType: KeyboardType, placeholder: String, onValueChange: (String) -> Unit) {
    val interactionSource = remember { MutableInteractionSource() }
    BasicTextField(
        value = value,
        onValueChange = onValueChange,
        singleLine = true,
        keyboardOptions = androidx.compose.foundation.text.KeyboardOptions(keyboardType = keyboardType, imeAction = ImeAction.Done),
        textStyle = TextStyle(color = Navy, fontSize = 12.sp, lineHeight = 18.sp, fontWeight = FontWeight.Bold, textAlign = TextAlign.Center),
        modifier = modifier.height(42.dp).padding(horizontal = 2.dp),
        decorationBox = { innerTextField ->
            OutlinedTextFieldDefaults.DecorationBox(
                value = value,
                innerTextField = innerTextField,
                enabled = true,
                singleLine = true,
                visualTransformation = VisualTransformation.None,
                interactionSource = interactionSource,
                placeholder = { Text(placeholder, color = Color(0xFF8A969C), fontSize = 10.sp) },
                contentPadding = PaddingValues(0.dp),
            )
        },
    )
}

@Composable
private fun ProductRow(product: ProductLine, tab: SalesTab, quantities: MutableMap<String, String>, rowIndex: Int, onChange: (String, String) -> Unit) {
    val rowColor = if (rowIndex % 2 == 0) Color.White else RowAlt
    Row(Modifier.fillMaxWidth().background(rowColor).padding(vertical = 5.dp, horizontal = 3.dp), verticalAlignment = Alignment.CenterVertically) {
        Column(Modifier.weight(1.55f).padding(start = 7.dp, end = 3.dp), horizontalAlignment = Alignment.End) {
            if (tab == SalesTab.TOTAL) {
                Text(product.name, color = Navy, fontSize = 13.sp, fontWeight = FontWeight.Bold, maxLines = 1, modifier = Modifier.fillMaxWidth(), textAlign = TextAlign.End)
            } else {
                Row(Modifier.fillMaxWidth(), verticalAlignment = Alignment.CenterVertically) {
                    Text(product.name, color = Navy, fontSize = 13.sp, fontWeight = FontWeight.Bold, maxLines = 1, modifier = Modifier.weight(1f), textAlign = TextAlign.End)
                    Text(priceLabel(product, tab), color = Color(0xFF687782), fontSize = 10.sp, fontWeight = FontWeight.Bold, maxLines = 1, modifier = Modifier.width(60.dp), textAlign = TextAlign.Start)
                }
            }
        }
        if (tab != SalesTab.RETAIL) QuantityCell(product.carton != null, quantities[key(product.code, "carton_wholesale")].orEmpty(), Modifier.weight(1f), allowDecimal = true) { onChange("carton_wholesale", it) }
        if (tab != SalesTab.RETAIL) QuantityCell(product.box != null, quantities[key(product.code, "box_wholesale")].orEmpty(), Modifier.weight(1f), allowDecimal = false) { onChange("box_wholesale", it) }
        if (tab != SalesTab.WHOLESALE) QuantityCell(product.carton != null, quantities[key(product.code, "carton_retail")].orEmpty(), Modifier.weight(1f), allowDecimal = true) { onChange("carton_retail", it) }
        if (tab != SalesTab.WHOLESALE) QuantityCell(product.box != null, quantities[key(product.code, "box_retail")].orEmpty(), Modifier.weight(1f), allowDecimal = false) { onChange("box_retail", it) }
    }
    HorizontalDivider(color = Color(0xFFD7E2E6))
}

@Composable
@OptIn(ExperimentalMaterial3Api::class)
private fun QuantityCell(enabled: Boolean, value: String, modifier: Modifier, allowDecimal: Boolean, onValueChange: (String) -> Unit) {
    val interactionSource = remember { MutableInteractionSource() }
    BasicTextField(
        value = value,
        onValueChange = onValueChange,
        enabled = enabled,
        singleLine = true,
        keyboardOptions = androidx.compose.foundation.text.KeyboardOptions(
            keyboardType = if (allowDecimal) KeyboardType.Decimal else KeyboardType.Number,
            imeAction = ImeAction.Done,
        ),
        textStyle = TextStyle(
            color = Navy,
            fontSize = 13.sp,
            lineHeight = 18.sp,
            fontWeight = FontWeight.Bold,
            textAlign = TextAlign.Center,
        ),
        modifier = modifier.height(42.dp).padding(horizontal = 2.dp),
        decorationBox = { innerTextField ->
            OutlinedTextFieldDefaults.DecorationBox(
                value = value,
                innerTextField = innerTextField,
                enabled = enabled,
                singleLine = true,
                visualTransformation = VisualTransformation.None,
                interactionSource = interactionSource,
                contentPadding = PaddingValues(0.dp),
            )
        },
    )
}

private fun priceLabel(product: ProductLine, tab: SalesTab): String {
    fun price(value: Double?): String = value?.let { formatMoney(it) } ?: "—"
    return when (tab) {
        SalesTab.WHOLESALE -> "${price(product.carton?.wholesalePrice)}/${price(product.box?.wholesalePrice)}"
        SalesTab.RETAIL -> "${price(product.carton?.retailPrice)}/${price(product.box?.retailPrice)}"
        SalesTab.TOTAL -> "ج${price(product.carton?.wholesalePrice)} م${price(product.carton?.retailPrice)}"
        SalesTab.SPECIAL -> ""
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
private fun sanitizeQuantity(value: String, allowDecimal: Boolean): String {
    if (!allowDecimal) return value.filter { it.isDigit() }
    val normalized = value.replace(',', '.')
    val result = StringBuilder()
    var decimalSeen = false
    normalized.forEach { char ->
        when {
            char.isDigit() -> result.append(char)
            char == '.' && !decimalSeen -> { result.append(char); decimalSeen = true }
        }
    }
    return result.toString()
}
private fun formatMoney(value: Double): String = String.format(Locale.US, "%,.0f", value)
private fun dateLabel(key: String): String = runCatching { SimpleDateFormat("dd/MM/yyyy", Locale.US).format(SimpleDateFormat("yyyyMMdd", Locale.US).parse(key)!!) }.getOrDefault(key)

private fun buildProducts(catalog: List<MaterialCatalog>): List<ProductLine> = catalog.groupBy { it.code }.map { (code, items) ->
    fun merged(unit: MaterialUnit): MaterialCatalog? {
        val matches = items.filter { it.unit == unit }
        if (matches.isEmpty()) return null
        val first = matches.first()
        return first.copy(
            name = items.firstOrNull { it.name.isNotBlank() }?.name ?: first.name,
            costPrice = matches.firstNotNullOfOrNull { it.costPrice },
            wholesalePrice = matches.firstNotNullOfOrNull { it.wholesalePrice },
            retailPrice = matches.firstNotNullOfOrNull { it.retailPrice }
        )
    }
    ProductLine(code, items.firstOrNull { it.name.isNotBlank() }?.name.orEmpty(), merged(MaterialUnit.CARTON), merged(MaterialUnit.BOX))
}.sortedBy { it.code }

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
private fun saveDay(p: android.content.SharedPreferences, day: String, q: Map<String, String>, t: Totals, specialSales: List<SpecialSaleDraft>) {
    p.edit()
        .putString("day_$day", q.entries.joinToString("\n") { "${it.key}=${it.value}" })
        .putString("special_$day", specialSales.joinToString("\n") { listOf(it.name, it.quantity, it.price).joinToString("|") })
        .putString("summary_$day", "${t.revenue}")
        .putStringSet("saved_days", p.getStringSet("saved_days", emptySet()).orEmpty() + day)
        .apply()
}
private fun loadDay(p: android.content.SharedPreferences, day: String): Map<String, String> = p.getString("day_$day", "").orEmpty().lineSequence().mapNotNull { line -> line.split('=', limit = 2).takeIf { it.size == 2 }?.let { it[0] to it[1] } }.toMap()
private fun loadSpecialSales(p: android.content.SharedPreferences, day: String): List<SpecialSaleDraft> = p.getString("special_$day", "").orEmpty().lineSequence().mapNotNull { line ->
    val parts = line.split('|')
    if (parts.size < 3) return@mapNotNull null
    SpecialSaleDraft().apply { name = parts[0]; quantity = parts[1]; price = parts[2] }
}.toList()
private fun calculateSpecialTotal(rows: List<SpecialSaleDraft>): Double = rows.sumOf { amount(it.quantity) * amount(it.price) }
private fun clearSalesHistory(p: android.content.SharedPreferences) {
    val editor = p.edit()
    p.all.keys.filter { it == "saved_days" || it.startsWith("day_") || it.startsWith("special_") || it.startsWith("summary_") }.forEach(editor::remove)
    editor.apply()
}

private fun exportXlsx(context: Context, uri: android.net.Uri, date: String, catalog: List<MaterialCatalog>, q: Map<String, String>, specialSales: List<SpecialSaleDraft>) {
    val products = buildProducts(catalog)
    val workbook = XSSFWorkbook()
    workbook.setForceFormulaRecalculation(true)
    val sheet = workbook.createSheet("مبيعات ${date.replace('/', '-')}")
    sheet.setRightToLeft(true)
    val headers = listOf("رمز المادة", "المادة", "جملة طرد", "جملة علبة", "مفرق طرد", "مفرق علبة", "قيمة الجملة", "قيمة المفرق", "إجمالي المبيعات", "خاص كمية", "خاص سعر", "إجمالي الخاص")
    val numberFormat = workbook.createDataFormat().getFormat("#,##0")
    fun rgb(hex: String): XSSFColor = XSSFColor(hex.chunked(2).map { it.toInt(16).toByte() }.toByteArray(), null)
    fun makeStyle(fill: XSSFColor, bold: Boolean = false, numeric: Boolean = false, whiteText: Boolean = false): XSSFCellStyle = workbook.createCellStyle().apply {
        setFillForegroundColor(fill)
        setFillPattern(FillPatternType.SOLID_FOREGROUND)
        setAlignment(HorizontalAlignment.CENTER)
        setVerticalAlignment(VerticalAlignment.CENTER)
        setBorderLeft(BorderStyle.THIN)
        setBorderRight(BorderStyle.THIN)
        setBorderTop(BorderStyle.THIN)
        setBorderBottom(BorderStyle.THIN)
        if (numeric) dataFormat = numberFormat
        val font = workbook.createFont().apply { this.bold = bold; fontHeightInPoints = if (bold) 11 else 10; color = if (whiteText) IndexedColors.WHITE.index else IndexedColors.BLACK.index }
        setFont(font)
    }
    val headerStyle = makeStyle(rgb("0B7A8C"), bold = true, whiteText = true)
    val whiteFill = rgb("FFFFFF")
    val neutralFill = rgb("E8F4F6")
    val totalFill = rgb("D5ECEF")
    val bodyStyle = makeStyle(whiteFill)
    val altStyle = makeStyle(neutralFill)
    val bodyNumberStyle = makeStyle(whiteFill, numeric = true)
    val altNumberStyle = makeStyle(neutralFill, numeric = true)
    val totalStyle = makeStyle(totalFill, bold = true, numeric = true)
    val header = sheet.createRow(0)
    header.heightInPoints = 26f
    headers.forEachIndexed { i, h -> header.createCell(i).apply { setCellValue(h); cellStyle = headerStyle } }
    var excelRow = 1
    products.forEach { p ->
        val quantities = listOf("carton_wholesale", "box_wholesale", "carton_retail", "box_retail").map { amount(q[key(p.code, it)]) }
        if (quantities.all { it == 0.0 }) return@forEach
        val row = sheet.createRow(excelRow++)
        val excelNumber = excelRow + 1
        row.createCell(0).setCellValue(p.code)
        row.createCell(1).setCellValue(p.name)
        // Keep the same stable zebra pattern as the in-app table: the first
        // visible data row is white, followed by the light teal row.
        val visibleRowIndex = excelRow - 2
        val isAlternateRow = visibleRowIndex % 2 == 1
        val textStyle = if (isAlternateRow) altStyle else bodyStyle
        val numericStyle = if (isAlternateRow) altNumberStyle else bodyNumberStyle
        row.getCell(0).cellStyle = textStyle
        row.createCell(1).apply { setCellValue(p.name); cellStyle = textStyle }
        quantities.forEachIndexed { i, value -> row.createCell(i + 2).apply { setCellValue(value); cellStyle = numericStyle } }
        val cartonWholesalePrice = p.carton?.wholesalePrice ?: 0.0
        val boxWholesalePrice = p.box?.wholesalePrice ?: 0.0
        val cartonRetailPrice = p.carton?.retailPrice ?: 0.0
        val boxRetailPrice = p.box?.retailPrice ?: 0.0
        row.createCell(6).apply { setCellFormula("C$excelNumber*${cartonWholesalePrice}+D$excelNumber*${boxWholesalePrice}"); cellStyle = numericStyle }
        row.createCell(7).apply { setCellFormula("E$excelNumber*${cartonRetailPrice}+F$excelNumber*${boxRetailPrice}"); cellStyle = numericStyle }
        row.createCell(8).apply { setCellFormula("G$excelNumber+H$excelNumber+L$excelNumber"); cellStyle = numericStyle }
        row.createCell(9).cellStyle = numericStyle
        row.createCell(10).cellStyle = numericStyle
        row.createCell(11).cellStyle = numericStyle
    }
    val filledSpecialSales = specialSales.filter { it.name.isNotBlank() || it.quantity.isNotBlank() || it.price.isNotBlank() }
    filledSpecialSales.forEachIndexed { index, sale ->
        val row = sheet.createRow(excelRow++)
        val rowNumber = excelRow + 1
        val visibleRowIndex = excelRow - 2
        val isAlternateRow = visibleRowIndex % 2 == 1
        val textStyle = if (isAlternateRow) altStyle else bodyStyle
        val numericStyle = if (isAlternateRow) altNumberStyle else bodyNumberStyle
        row.createCell(0).apply { setCellValue(""); cellStyle = textStyle }
        row.createCell(1).apply { setCellValue(sale.name); cellStyle = textStyle }
        for (i in 2..7) row.createCell(i).apply { setCellValue(0.0); cellStyle = numericStyle }
        row.createCell(8).apply { setCellFormula("L$rowNumber"); cellStyle = numericStyle }
        row.createCell(9).apply { setCellValue(amount(sale.quantity)); cellStyle = numericStyle }
        row.createCell(10).apply { setCellValue(amount(sale.price)); cellStyle = numericStyle }
        row.createCell(11).apply { setCellFormula("J$rowNumber*K$rowNumber"); cellStyle = numericStyle }
    }
    val total = sheet.createRow(excelRow)
    total.createCell(0).apply { setCellValue("الإجمالي"); cellStyle = totalStyle }
    total.createCell(1).cellStyle = totalStyle
    val lastDataRow = excelRow
    for (i in 2..11) total.createCell(i).apply {
        val column = ('A'.code + i).toChar()
        setCellFormula(if (lastDataRow >= 2) "SUM(${column}2:${column}$lastDataRow)" else "0")
        cellStyle = totalStyle
    }
    sheet.createFreezePane(0, 1)
    sheet.setAutoFilter(CellRangeAddress(0, excelRow, 0, 11))
    listOf(14, 28, 12, 12, 12, 12, 16, 16, 18, 14, 14, 16).forEachIndexed { i, width -> sheet.setColumnWidth(i, width * 256) }
    val bytes = ByteArrayOutputStream().use { output ->
        workbook.write(output)
        output.toByteArray()
    }
    workbook.close()
    require(bytes.isNotEmpty()) { "تعذر إنشاء محتوى ملف Excel" }
    context.contentResolver.openOutputStream(uri)?.use { output ->
        output.write(bytes)
        output.flush()
    } ?: error("تعذر إنشاء ملف Excel")
}

private fun quantitiesTotal(products: List<ProductLine>, q: Map<String, String>, index: Int): Double = products.sumOf { p -> amount(q[key(p.code, listOf("carton_wholesale", "box_wholesale", "carton_retail", "box_retail")[index])]) }

private fun exportHistoryXlsx(context: Context, uri: android.net.Uri, preferences: android.content.SharedPreferences) {
    val workbook = XSSFWorkbook()
    workbook.setForceFormulaRecalculation(true)
    val sheet = workbook.createSheet("سجل المبيعات")
    sheet.setRightToLeft(true)
    val numberFormat = workbook.createDataFormat().getFormat("#,##0")
    fun rgb(hex: String): XSSFColor = XSSFColor(hex.chunked(2).map { it.toInt(16).toByte() }.toByteArray(), null)
    fun makeStyle(fill: XSSFColor, bold: Boolean = false, numeric: Boolean = false, whiteText: Boolean = false): XSSFCellStyle = workbook.createCellStyle().apply {
        setFillForegroundColor(fill)
        setFillPattern(FillPatternType.SOLID_FOREGROUND)
        setAlignment(HorizontalAlignment.CENTER)
        setVerticalAlignment(VerticalAlignment.CENTER)
        setBorderLeft(BorderStyle.THIN)
        setBorderRight(BorderStyle.THIN)
        setBorderTop(BorderStyle.THIN)
        setBorderBottom(BorderStyle.THIN)
        if (numeric) dataFormat = numberFormat
        val font = workbook.createFont().apply { this.bold = bold; fontHeightInPoints = if (bold) 11 else 10; color = if (whiteText) IndexedColors.WHITE.index else IndexedColors.BLACK.index }
        setFont(font)
    }
    val headerStyle = makeStyle(rgb("0B7A8C"), bold = true, whiteText = true)
    val bodyStyle = makeStyle(rgb("FFFFFF"))
    val altStyle = makeStyle(rgb("E8F4F6"))
    val bodyNumberStyle = makeStyle(rgb("FFFFFF"), numeric = true)
    val altNumberStyle = makeStyle(rgb("E8F4F6"), numeric = true)
    val totalStyle = makeStyle(rgb("D5ECEF"), bold = true, numeric = true)
    val header = sheet.createRow(0)
    header.heightInPoints = 26f
    listOf("التاريخ", "إجمالي المبيعات").forEachIndexed { index, title ->
        header.createCell(index).apply { setCellValue(title); cellStyle = headerStyle }
    }
    val days = preferences.getStringSet("saved_days", emptySet()).orEmpty().toList().sorted()
    days.forEachIndexed { index, day ->
        val row = sheet.createRow(index + 1)
        val textStyle = if (index % 2 == 0) bodyStyle else altStyle
        val numericStyle = if (index % 2 == 0) bodyNumberStyle else altNumberStyle
        row.createCell(0).apply { setCellValue(dateLabel(day)); cellStyle = textStyle }
        row.createCell(1).apply {
            setCellValue(preferences.getString("summary_$day", "0").orEmpty().split('|').firstOrNull()?.toDoubleOrNull() ?: 0.0)
            cellStyle = numericStyle
        }
    }
    val totalRow = sheet.createRow(days.size + 1)
    totalRow.createCell(0).apply { setCellValue("الإجمالي"); cellStyle = totalStyle }
    totalRow.createCell(1).apply {
        setCellFormula(if (days.isEmpty()) "0" else "SUM(B2:B${days.size + 1})")
        cellStyle = totalStyle
    }
    sheet.createFreezePane(0, 1)
    sheet.setAutoFilter(CellRangeAddress(0, days.size + 1, 0, 1))
    sheet.setColumnWidth(0, 18 * 256)
    sheet.setColumnWidth(1, 22 * 256)
    val bytes = ByteArrayOutputStream().use { output ->
        workbook.write(output)
        output.toByteArray()
    }
    workbook.close()
    require(bytes.isNotEmpty()) { "تعذر إنشاء محتوى ملف سجل Excel" }
    context.contentResolver.openOutputStream(uri)?.use { output ->
        output.write(bytes)
        output.flush()
    } ?: error("تعذر إنشاء ملف سجل Excel")
}
