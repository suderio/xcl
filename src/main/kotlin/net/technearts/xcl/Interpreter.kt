package net.technearts.xcl

import bsh.Interpreter
import java.util.regex.Pattern
import java.util.stream.Collectors

class SheeterInterpreter : Interpreter() {
    val cells = mutableListOf<XCLCell<*>>()
    private val columnPattern = Pattern.compile("\\$\\d+")

    fun <T> allParametersKnown(currentCell: XCLCell<T>, columns: List<String>): Boolean {
        val currentColumn = currentCell.address.column
        return columns.filterIndexed { index, _ -> index < currentColumn }.flatMap { it.referencedColumns() }
            .none { column -> column > currentColumn }
    }

    private fun String.referencedColumns(): Set<Int> {
        return columnPattern.matcher(this).results().map { r -> r.group().replace("$", "") }.map(String::toInt)
            .collect(Collectors.toSet())
    }

}