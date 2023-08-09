package net.technearts.xcl

import bsh.Interpreter
import java.io.Closeable
import java.io.File
import java.util.function.Consumer
import java.util.regex.Pattern

fun repl(input: CellProducer, output: CellConsumer, columns: List<String>) {
    val engine = Interpreter()
    output.use { writer ->
        input.use { reader ->
            reader.asSequence().map { xclCell -> columns.script(xclCell, engine) }.forEach { writer.accept(it) }
        }
    }
}

fun <T> List<String>.script(cell: XCLCell<T>, engine: Interpreter): XCLCell<T> {
    engine.set("__cell__${cell.row}__${cell.col}", cell.value)
    return if (this.isEmpty()) {
        cell
    } else {
        val expr = this[cell.col].replace("$", "__cell__${cell.row}__")
        engine.eval("current = $expr")
        @Suppress("UNCHECKED_CAST")
        val value: T = engine.get("current") as T
        XCLCell(cell.row, cell.col, cell.type, value)
    }
}

fun inExcelStream(input: File?, tab: String, startCell: String, endCell: String): CellProducer {
    return XSSFWorkbookReader(
        input?.inputStream() ?: System.`in`,
        tab,
        startCell.address(),
        endCell.address()
    )
}

fun inCSVStream(input: File?, delimiter: String, separator: String): CellProducer {
    return CSVReader(input?.inputStream() ?: System.`in`, builder(delimiter, separator).build())
}

fun outCSVStream(output: File?, delimiter: String, separator: String): CellConsumer {
    return CSVWriter(output?.outputStream() ?: System.out, builder(delimiter, separator).build())
}

fun outExcelStream(output: File?, tab: String, startCell: String, endCell: String): CellConsumer {
    return XSSFWorkbookWriter(output?.outputStream() ?: System.out, tab, startCell.address(), endCell.address())
}

fun <T> pair(i: Iterable<T>) = Pair(i.elementAt(0), i.elementAt(1))

val r1c1Pattern: Pattern = Pattern.compile("R\\d+C\\d+", Pattern.CASE_INSENSITIVE)
fun String.address(): Pair<Int, Int>? {
    if (this.isEmpty()) return null
    if (!r1c1Pattern.matcher(this).matches()) return null
    return pair(this.split('r', 'c', ignoreCase = true, limit = 0).map(String::toInt))
}

data class XCLCell<T>(val row: Int, val col: Int, val type: XCLCellType, val value: T)

enum class XCLCellType { NUMBER, STRING, DATE, FORMULA, ERROR, BOOLEAN, EMPTY }

interface CellConsumer : Consumer<XCLCell<*>>, Closeable

interface CellProducer : Sequence<XCLCell<*>>, Closeable


