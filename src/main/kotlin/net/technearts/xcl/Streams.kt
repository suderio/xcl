package net.technearts.xcl

import java.io.Closeable
import java.io.File
import java.util.function.Consumer

fun repl(input: CellProducer, output: CellConsumer, columns: List<String>, show: Boolean = false) {
    val engine = SheeterInterpreter()
    engine.showResults = show
    output.use { writer ->
        input.use { reader ->
            reader.asSequence()
                .flatMap { xclCell -> engine.eval(xclCell, columns) }
                .forEach { writer.accept(it) }
        }
    }
}

fun inExcelStream(input: File?, tab: String, startCell: Address = Address(0,0), endCell: Address? = null, removeHeaders: Boolean = false): CellProducer {
    return XSSFWorkbookReader(input?.inputStream() ?: System.`in`, tab, startCell, endCell, removeHeaders)
}

fun inCSVStream(input: File?, delimiter: String, separator: String, startCell: Address = Address(0,0), endCell: Address? = null, removeHeaders: Boolean = false): CellProducer {
    return CSVReader(input?.inputStream() ?: System.`in`, builder(delimiter, separator).build(), startCell, endCell, removeHeaders)
}

fun outCSVStream(output: File?, delimiter: String, separator: String, headers: List<String> = emptyList()): CellConsumer {
    return CSVWriter(output?.outputStream() ?: System.out, builder(delimiter, separator).build(), headers)
}

fun outExcelStream(output: File?, tab: String, startCell: Address = Address(0,0), headers: List<String> = emptyList()): CellConsumer {
    return XSSFWorkbookWriter(output?.outputStream() ?: System.out, tab, startCell, headers)
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


