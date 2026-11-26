package com.example.app.excel

/** A single sellable unit from the user's material price catalog. */
data class MaterialCatalog(
    val code: String,
    val name: String,
    val unit: MaterialUnit,
    val costPrice: Double?,
    val wholesalePrice: Double?,
    val retailPrice: Double?,
)

enum class MaterialUnit(val label: String) {
    CARTON("طرد"),
    BOX("علبة"),
    OTHER("");

    companion object {
        fun from(value: String): MaterialUnit = when (value.trim()) {
            "طرد" -> CARTON
            "علبة" -> BOX
            else -> OTHER
        }
    }
}
