package net.technearts.xcl

import io.quarkus.arc.log.LoggerName
import org.dhatim.fastexcel.VisibilityState
import org.dhatim.fastexcel.Workbook
import org.dhatim.fastexcel.reader.CellType.*
import org.dhatim.fastexcel.reader.ReadableWorkbook
import org.jboss.logging.Logger
import java.io.*
import java.util.function.Consumer
import java.util.stream.Stream
import kotlin.streams.asSequence

const val FS = '\u00F1'
const val LF = '\u000A'

fun repl(input: CellProducer, output: CellConsumer) {
    output.use { writer ->
        input.use { reader ->
            reader.asSequence().forEach { writer.accept(it) }
        }
    }
}

fun inExcelStream(input: File?, tab: String): CellProducer {
    return WorkbookReader(input?.inputStream() ?: System.`in`, tab)
}

fun inCSVStream(input: File?): CellProducer {
    return CSVReader(input?.inputStream() ?: System.`in`)
}

fun outCSVStream(output: File?): CellConsumer {
    return CSVWriter(output?.outputStream() ?: System.out)
}

fun outExcelStream(output: File?, tab: String): CellConsumer {
    return WorkbookWriter(output?.outputStream() ?: System.out, tab)
}

data class XCLCell<T>(val row: Int, val col: Int, val type: XCLCellType, val value: T, val last: Boolean = false)
enum class XCLCellType { NUMBER, STRING, DATE, FORMULA, ERROR, BOOLEAN, EMPTY }

interface CellConsumer : Consumer<XCLCell<*>>, Closeable

interface CellProducer : Sequence<XCLCell<*>>, Closeable

class CSVReader(stream: InputStream) : BufferedReader(InputStreamReader(stream)), CellProducer {
    override fun iterator(): Iterator<XCLCell<*>> {
        return this.lines().asSequence().flatMapIndexed { row: Int, line: String ->
            var col = 0
            line.splitToSequence(",").map { s: String ->
                XCLCell(row, col++, XCLCellType.STRING, s)
                // TODO Parser heurístico para tentar descobrir o tipo
            }
        }.iterator()
    }
}

class CSVWriter(stream: OutputStream) : FilterWriter(BufferedWriter(OutputStreamWriter(stream))), CellConsumer {
    override fun accept(cell: XCLCell<*>) {
        this.write(cell.value.toString())
        if (cell.last) (this.out as BufferedWriter).newLine()
        else this.write(",")
    }
}

class WorkbookReader(stream: InputStream, private val tab: String) : BufferedInputStream(stream), CellProducer {
    private val workbook = ReadableWorkbook(this)
    private val sheet = workbook.findSheet(tab)

    @LoggerName("xcl")
    private lateinit var log: Logger

    override fun iterator(): Iterator<XCLCell<*>> {
        return if (sheet.isPresent) {
            sheet.get().openStream().flatMap { row ->
                row.stream().map { cell ->
                    val r = cell.address.row
                    val c = cell.address.column
                    val last = c == row.cellCount - 1
                    when (cell.type) {
                        NUMBER -> XCLCell(r, c, XCLCellType.NUMBER, cell.asNumber(), last)
                        STRING -> XCLCell(r, c, XCLCellType.STRING, cell.asString(), last)
                        FORMULA -> XCLCell(r, c, XCLCellType.FORMULA, cell.formula, last)
                        ERROR -> XCLCell(r, c, XCLCellType.ERROR, cell.rawValue, last)
                        BOOLEAN -> XCLCell(r, c, XCLCellType.BOOLEAN, cell.asBoolean(), last)
                        EMPTY -> XCLCell(r, c, XCLCellType.EMPTY, "", last)
                        else -> XCLCell(r, c, XCLCellType.ERROR, "", last)
                    }
                }
            }
        } else {
            log.warn("Excel Tab $tab not found")
            Stream.empty()
        }.iterator()
    }
}

class WorkbookWriter(stream: OutputStream, tab: String) : Workbook(stream, "xcl", "1.0"), CellConsumer {
    private val sheet = this.newWorksheet(tab)

    init {
        sheet.keepInActiveTab()
        sheet.visibilityState = VisibilityState.VISIBLE;
    }
    override fun accept(cell: XCLCell<*>) {
        // TODO tratar pelo tipo da cell
        sheet.value(cell.row, cell.col, cell.value.toString())
    }
}