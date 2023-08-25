package net.technearts.xcl

import bsh.Interpreter
import net.technearts.xcl.XCLCellType.*
import java.time.LocalDate
import java.time.LocalDateTime
import java.util.*
import java.util.regex.Pattern

class SheeterInterpreter : Interpreter() {
    private val columnPattern = Pattern.compile("\\$(\\d+)")

    private fun <T> allParametersKnown(currentCell: XCLCell<T>, columns: List<String>): Boolean {
        val currentColumn = currentCell.address.column
        return columns.map { biggestColumnReference(it) }
            .none { column -> column > currentColumn }
    }

    private fun biggestColumnReference(column: String): Int {
        return columnPattern.matcher(column).results()
            .map { r -> r.group().replace("$", "") }
            .map(String::toInt)
            .reduce(Math::max)
            .get()
    }

    fun <T> eval(cell: XCLCell<T>, columns: List<String>): Sequence<XCLCell<*>> {
        // __cell__$row__$col = original cell value
        this.set("__cell__${cell.address.row}__${cell.address.column}", cell.value)

        // if columns is empty, no need to do anything else
        if (columns.isEmpty())
            return sequenceOf(cell)

        // if current cell column does not have enough info to eval all columns, do nothing
        if (!allParametersKnown(cell, columns)) {
            return emptySequence()
        }

        // replace every column reference with the previous cell values and evaluate the expression
        return columns.asSequence()
            .map { s ->
                columnPattern.matcher(s)
                    .replaceAll { match -> "__cell__${cell.address.row}__${match.group().replace("$", "")}" }
            }
            .onEach { this.eval("current = $it") }
            .mapIndexed { i, _ -> createCellFromCurrent(cell.address.row, i)
            }

    }

    private fun createCellFromCurrent(row: Int, column: Int): XCLCell<*> {
        val result = this.get("current")
        return XCLCell(
            Address(row, column), when (result) {
                is Boolean -> BOOLEAN
                is Number -> NUMBER
                is Date, is LocalDate, is LocalDateTime -> DATE
                else -> STRING
            }, result
        )
    }
}