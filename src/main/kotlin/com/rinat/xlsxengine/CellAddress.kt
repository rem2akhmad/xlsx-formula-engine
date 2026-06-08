package com.rinat.xlsxengine

/** 1-based row/column address in a worksheet. */
data class CellAddress(val row: Int, val column: Int) {
    init {
        require(row >= 1) { "Row must be >= 1" }
        require(column >= 1) { "Column must be >= 1" }
    }

    fun toA1(): String = columnToName(column) + row

    companion object {
        // Excel columns are in range A..XFD (1..3 letters).
        private val A1_REGEX = Regex("^\\$?([A-Za-z]{1,3})\\$?([0-9]{1,7})$")
        private val R1C1_REGEX = Regex("^R([0-9]{1,7})C([0-9]{1,7})$", RegexOption.IGNORE_CASE)

        fun parse(a1: String): CellAddress {
            val normalized = a1.trim()
            val r1c1 = R1C1_REGEX.matchEntire(normalized)
            if (r1c1 != null) {
                return CellAddress(
                    row = r1c1.groupValues[1].toInt(),
                    column = r1c1.groupValues[2].toInt()
                )
            }

            val match = A1_REGEX.matchEntire(normalized)
                ?: error("Unsupported cell reference: '$a1'")

            val column = nameToColumn(match.groupValues[1])
            val row = match.groupValues[2].toInt()
            return CellAddress(row = row, column = column)
        }

        fun nameToColumn(name: String): Int {
            require(name.isNotBlank()) { "Column name must not be blank" }
            var value = 0
            for (ch in name.uppercase()) {
                require(ch in 'A'..'Z') { "Invalid column name: $name" }
                value = value * 26 + (ch.code - 'A'.code + 1)
            }
            return value
        }

        fun columnToName(columnIndex: Int): String {
            require(columnIndex >= 1) { "Column index must be >= 1" }
            var idx = columnIndex
            val buffer = StringBuilder()
            while (idx > 0) {
                idx -= 1
                buffer.append(('A'.code + (idx % 26)).toChar())
                idx /= 26
            }
            return buffer.reverse().toString()
        }
    }
}
