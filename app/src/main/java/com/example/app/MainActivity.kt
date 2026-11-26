package com.example.app

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
import androidx.compose.material3.Button
import androidx.compose.material3.ButtonDefaults
import androidx.compose.material3.Card
import androidx.compose.material3.CardDefaults
import androidx.compose.material3.HorizontalDivider
import androidx.compose.material3.MaterialTheme
import androidx.compose.material3.OutlinedButton
import androidx.compose.material3.OutlinedTextField
import androidx.compose.material3.Surface
import androidx.compose.material3.Text
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
import java.text.SimpleDateFormat
import java.util.Date
import java.util.Locale

private val Teal = Color(0xFF1196A7)
private val Navy = Color(0xFF10283D)
private val Page = Color(0xFFF7F3EA)
private val RowAlt = Color(0xFFF4F7F8)
private val Green = Color(0xFF2E7D50)

private data class ProductLine(
    val code: String,
    val name: String,
    val carton: MaterialCatalog?,
    val box: MaterialCatalog?,
)

class MainActivity : ComponentActivity() {
    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)
        setContent { SalesMovementApp() }
    }
}

@Composable
private fun SalesMovementApp() {
    val context = androidx.compose.ui.platform.LocalContext.current
    val preferences = remember { context.getSharedPreferences("sales_movement_data", 0) }
    var catalog by remember { mutableStateOf(loadCatalog(preferences)) }
    val quantities = remember { mutableStateMapOf<String, String>() }
    var message by remember { mutableStateOf("") }
    val date = remember { SimpleDateFormat("dd/MM/yyyy", Locale.US).format(Date()) }
    val launcher = rememberLauncherForActivityResult(ActivityResultContracts.OpenDocument()) { uri ->
        if (uri == null) return@rememberLauncherForActivityResult
        runCatching {
            val items = MaterialCatalogReader(context.contentResolver).read(uri)
            saveCatalog(preferences, items)
            catalog = items
            message = "تم تحميل ${items.map { it.code }.distinct().size} مادة"
        }.onFailure { error -> message = error.message ?: "فشل قراءة دليل المواد" }
    }

    val products = catalog.groupBy { it.code }.map { (code, items) ->
        ProductLine(
            code = code,
            name = items.first().name,
            carton = items.firstOrNull { it.unit == MaterialUnit.CARTON },
            box = items.firstOrNull { it.unit == MaterialUnit.BOX },
        )
    }.sortedBy { it.code }
    val totalSales = products.sumOf { product ->
        amount(quantities[key(product.code, "carton_wholesale")]) * (product.carton?.wholesalePrice ?: 0.0) +
            amount(quantities[key(product.code, "box_wholesale")]) * (product.box?.wholesalePrice ?: 0.0) +
            amount(quantities[key(product.code, "carton_retail")]) * (product.carton?.retailPrice ?: 0.0) +
            amount(quantities[key(product.code, "box_retail")]) * (product.box?.retailPrice ?: 0.0)
    }

    MaterialTheme {
        Surface(modifier = Modifier.fillMaxSize(), color = Page) {
            Column(modifier = Modifier.fillMaxSize()) {
                Column(
                    modifier = Modifier.fillMaxWidth().background(Teal).padding(horizontal = 18.dp, vertical = 18.dp),
                    horizontalAlignment = Alignment.End,
                ) {
                    Text("مبيعات المندوب", color = Color.White, fontSize = 26.sp, fontWeight = FontWeight.Bold)
                    Text("مبيعات اليوم — $date", color = Color.White, fontSize = 15.sp)
                }
                Column(modifier = Modifier.padding(16.dp)) {
                    Row(modifier = Modifier.fillMaxWidth(), horizontalArrangement = Arrangement.spacedBy(8.dp)) {
                        Button(
                            onClick = { launcher.launch(arrayOf("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "application/vnd.ms-excel")) },
                            colors = ButtonDefaults.buttonColors(containerColor = Teal),
                            modifier = Modifier.weight(1f),
                        ) { Text("رفع دليل المواد") }
                        OutlinedButton(onClick = { quantities.clear(); message = "تم تصفير مبيعات اليوم" }, modifier = Modifier.weight(1f)) { Text("تصفير اليوم") }
                    }
                    Spacer(Modifier.height(10.dp))
                    Text("المواد: ${products.size}    إجمالي المبيعات: ${formatMoney(totalSales)}", color = Navy, fontSize = 17.sp, fontWeight = FontWeight.Bold)
                    if (message.isNotBlank()) Text(message, color = Green, fontSize = 14.sp, modifier = Modifier.padding(top = 4.dp))
                }
                if (products.isEmpty()) {
                    Card(modifier = Modifier.padding(16.dp).fillMaxWidth(), colors = CardDefaults.cardColors(containerColor = Color.White), shape = RoundedCornerShape(16.dp)) {
                        Column(modifier = Modifier.padding(24.dp), horizontalAlignment = Alignment.End) {
                            Text("لم يتم تحميل دليل المواد بعد", color = Navy, fontSize = 20.sp, fontWeight = FontWeight.Bold)
                            Spacer(Modifier.height(8.dp))
                            Text("ارفع ملف Excel الذي يحتوي على رمز المادة والاسم والوحدة والأسعار.", color = Navy, fontSize = 15.sp, textAlign = TextAlign.Right)
                        }
                    }
                } else {
                    SalesHeader()
                    LazyColumn(modifier = Modifier.fillMaxSize().padding(horizontal = 12.dp)) {
                        items(products) { product ->
                            ProductRow(product, quantities) { field, value ->
                                quantities[key(product.code, field)] = value.filter { it.isDigit() || it == '.' }
                                saveSales(preferences, quantities)
                            }
                        }
                        item { Spacer(Modifier.height(24.dp)) }
                    }
                }
            }
        }
    }
}

@Composable
private fun SalesHeader() {
    Row(modifier = Modifier.fillMaxWidth().background(Navy).padding(vertical = 10.dp, horizontal = 6.dp), verticalAlignment = Alignment.CenterVertically) {
        Text("المادة", color = Color.White, fontWeight = FontWeight.Bold, modifier = Modifier.weight(1.55f), textAlign = TextAlign.Center)
        Text("جملة\nطرد", color = Color.White, fontWeight = FontWeight.Bold, modifier = Modifier.weight(1f), textAlign = TextAlign.Center)
        Text("جملة\nعلبة", color = Color.White, fontWeight = FontWeight.Bold, modifier = Modifier.weight(1f), textAlign = TextAlign.Center)
        Text("مفرق\nطرد", color = Color.White, fontWeight = FontWeight.Bold, modifier = Modifier.weight(1f), textAlign = TextAlign.Center)
        Text("مفرق\nعلبة", color = Color.White, fontWeight = FontWeight.Bold, modifier = Modifier.weight(1f), textAlign = TextAlign.Center)
    }
}

@Composable
private fun ProductRow(product: ProductLine, quantities: MutableMap<String, String>, onChange: (String, String) -> Unit) {
    val rowColor = if (product.code.hashCode() % 2 == 0) Color.White else RowAlt
    Row(modifier = Modifier.fillMaxWidth().background(rowColor).padding(vertical = 8.dp, horizontal = 4.dp), verticalAlignment = Alignment.CenterVertically) {
        Column(modifier = Modifier.weight(1.55f).padding(horizontal = 3.dp), horizontalAlignment = Alignment.End) {
            Text(product.name, color = Navy, fontSize = 14.sp, fontWeight = FontWeight.Bold, maxLines = 1)
            Text(product.code, color = Color(0xFF687782), fontSize = 11.sp)
        }
        QuantityCell(product.carton?.wholesalePrice != null, quantities[key(product.code, "carton_wholesale")].orEmpty(), Modifier.weight(1f)) { onChange("carton_wholesale", it) }
        QuantityCell(product.box?.wholesalePrice != null, quantities[key(product.code, "box_wholesale")].orEmpty(), Modifier.weight(1f)) { onChange("box_wholesale", it) }
        QuantityCell(product.carton?.retailPrice != null, quantities[key(product.code, "carton_retail")].orEmpty(), Modifier.weight(1f)) { onChange("carton_retail", it) }
        QuantityCell(product.box?.retailPrice != null, quantities[key(product.code, "box_retail")].orEmpty(), Modifier.weight(1f)) { onChange("box_retail", it) }
    }
    HorizontalDivider(color = Color(0xFFD7E2E6))
}

@Composable
private fun QuantityCell(enabled: Boolean, value: String, modifier: Modifier, onValueChange: (String) -> Unit) {
    OutlinedTextField(
        value = value,
        onValueChange = onValueChange,
        enabled = enabled,
        singleLine = true,
        textStyle = TextStyle(fontSize = 13.sp, textAlign = TextAlign.Center),
        modifier = modifier.padding(horizontal = 2.dp),
        placeholder = { Text("0", fontSize = 12.sp) },
    )
}

private fun key(code: String, field: String) = "$code:$field"
private fun amount(value: String?): Double = value?.toDoubleOrNull() ?: 0.0
private fun formatMoney(value: Double): String = String.format(Locale.US, "%,.2f", value)

private fun saveCatalog(preferences: android.content.SharedPreferences, items: List<MaterialCatalog>) {
    preferences.edit().putString("catalog", items.joinToString("\n") { listOf(it.code, it.name, it.unit.label, it.costPrice ?: "", it.wholesalePrice ?: "", it.retailPrice ?: "").joinToString("|") }).apply()
}

private fun loadCatalog(preferences: android.content.SharedPreferences): List<MaterialCatalog> = preferences.getString("catalog", "").orEmpty().lineSequence().mapNotNull { line ->
    val parts = line.split('|')
    if (parts.size < 6) return@mapNotNull null
    val unit = MaterialUnit.from(parts[2])
    if (unit == MaterialUnit.OTHER) return@mapNotNull null
    MaterialCatalog(parts[0], parts[1], unit, parts[3].toDoubleOrNull(), parts[4].toDoubleOrNull(), parts[5].toDoubleOrNull())
}.toList()

private fun saveSales(preferences: android.content.SharedPreferences, quantities: Map<String, String>) {
    preferences.edit().putString("today_sales", quantities.entries.joinToString("\n") { "${it.key}=${it.value}" }).apply()
}
