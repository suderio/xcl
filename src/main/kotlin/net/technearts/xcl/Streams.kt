package net.technearts.xcl

import org.dhatim.fastexcel.reader.CellType
import org.dhatim.fastexcel.reader.ReadableWorkbook
import java.io.*
import java.util.function.Consumer
import java.util.stream.Stream
import kotlin.streams.asSequence

const val FS = '\u00F1'
const val LF = '\u000A'

fun repl(input: Sequence<XCLCell<*>>, output: CellConsumer) {
    output.use { w ->
        input.forEach { w.accept(it) }
    }
}

data class XCLCell<T>(val row: Int, val col: Int, val type: XCLCellType, val value: T)
enum class XCLCellType { NUMBER, STRING, DATE, FORMULA, ERROR, BOOLEAN, EMPTY }

interface CellConsumer : Consumer<XCLCell<*>>, Closeable
fun inStream(input: File?, tab: String): Sequence<XCLCell<*>> {
    return WorkbookStream(input?.inputStream() ?: System.`in`, tab).stream().asSequence()
}

fun outStream(output: File?): CellConsumer {
    return CSVWrite(output?.outputStream() ?: System.out)
}

class CSVWrite(stream: OutputStream) : FilterWriter(BufferedWriter(OutputStreamWriter(stream))), CellConsumer {
    override fun accept(cell: XCLCell<*>) {
        if (cell.col == 0) {
            if (cell.row != 0)
                this.append('\u0008').write("\n")
        }
        this.append(cell.value.toString()).write(",")
    }

}
class WorkbookStream(stream: InputStream, tab: String) {
    private val workbook = ReadableWorkbook(BufferedInputStream(stream))
    private val sheet = workbook.findSheet(tab)

    fun stream(): Stream<XCLCell<*>> = if (sheet.isPresent) {
        sheet.get().openStream().flatMap { rows ->
            rows.stream().map { cell ->
                when (cell.type) {
                    CellType.NUMBER -> XCLCell(
                        cell.address.row,
                        cell.address.column,
                        XCLCellType.NUMBER,
                        cell.asNumber()
                    )

                    CellType.STRING -> XCLCell(
                        cell.address.row,
                        cell.address.column,
                        XCLCellType.STRING,
                        cell.asString()
                    )

                    CellType.FORMULA -> XCLCell(
                        cell.address.row,
                        cell.address.column,
                        XCLCellType.FORMULA,
                        cell.formula
                    )

                    CellType.ERROR -> XCLCell(cell.address.row, cell.address.column, XCLCellType.ERROR, cell.rawValue)
                    CellType.BOOLEAN -> XCLCell(
                        cell.address.row,
                        cell.address.column,
                        XCLCellType.BOOLEAN,
                        cell.asBoolean()
                    )

                    CellType.EMPTY -> XCLCell(cell.address.row, cell.address.column, XCLCellType.EMPTY, "")
                    else -> XCLCell(cell.address.row, cell.address.column, XCLCellType.ERROR, "")
                }
            }
        }
    } else {
        Stream.empty()
    }
}
