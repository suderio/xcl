package net.technearts.xcl

import java.io.Closeable
import java.io.File
import java.util.function.Consumer

fun repl(input: CellProducer, output: CellConsumer, columns: List<String>) {
    val engine = SheeterInterpreter()
    output.use { writer ->
        input.use { reader ->
            reader.asSequence().flatMap { xclCell -> columns.script(xclCell, engine) }.forEach { writer.accept(it) }
        }
    }
}

fun <T> List<String>.script(cell: XCLCell<T>, engine: SheeterInterpreter): Sequence<XCLCell<T>> {
    // __cell__$row__$col = original cell value
    engine.set("__cell__${cell.address.row}__${cell.address.column}", cell.value)
    engine.cells += cell
    return if (this.isEmpty() || !engine.allParametersKnown(cell, this)) {
        sequenceOf(cell)
    } else {
        // changes every $n in the current 'column' with __cell__$row__n and sets it as 'current'
        val expr = this[cell.address.column].replace("$", "__cell__${cell.address.row}__")
        engine.eval("current = $expr")
        // eval 'current' and sets as the returning cell value
        @Suppress("UNCHECKED_CAST") val value: T = engine.get("current") as T
        sequenceOf( XCLCell(cell.address, cell.type, value))
    }
}

fun inExcelStream(input: File?, tab: String, startCell: Address, endCell: Address): CellProducer {
    return XSSFWorkbookReader(input?.inputStream() ?: System.`in`, tab, startCell, endCell)
}

fun inCSVStream(input: File?, delimiter: String, separator: String, startCell: Address, endCell: Address): CellProducer {
    return CSVReader(input?.inputStream() ?: System.`in`, builder(delimiter, separator).build(), startCell, endCell)
}

fun outCSVStream(output: File?, delimiter: String, separator: String): CellConsumer {
    return CSVWriter(output?.outputStream() ?: System.out, builder(delimiter, separator).build())
}

fun outExcelStream(output: File?, tab: String, startCell: Address): CellConsumer {
    return XSSFWorkbookWriter(output?.outputStream() ?: System.out, tab, startCell)
}

data class Address(
    val row: Int, val column: Int
) {
    override fun toString(): String = "($row, $column)"

}

data class XCLCell<T>(val address: Address, val type: XCLCellType, val value: T)

enum class XCLCellType { NUMBER, STRING, DATE, FORMULA, ERROR, BOOLEAN, EMPTY }

interface CellConsumer : Consumer<XCLCell<*>>, Closeable

interface CellProducer : Sequence<XCLCell<*>>, Closeable


