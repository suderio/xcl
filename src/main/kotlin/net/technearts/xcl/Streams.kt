package net.technearts.xcl

import bsh.Interpreter
import net.technearts.xcl.XCLCellType.*
import org.apache.poi.ss.usermodel.CellType
import org.apache.poi.ss.usermodel.DateUtil
import org.apache.poi.ss.usermodel.Row
import org.apache.poi.xssf.streaming.SXSSFCell
import org.apache.poi.xssf.streaming.SXSSFRow
import org.apache.poi.xssf.streaming.SXSSFSheet
import org.apache.poi.xssf.streaming.SXSSFWorkbook
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.*
import java.time.LocalDate
import java.util.function.Consumer
import kotlin.streams.asSequence


const val FS = '\u00F1'
const val LF = '\u000A'

fun repl(input: CellProducer, output: CellConsumer, columns: List<String>) {
    val engine = Interpreter()
    output.use { writer ->
        input.use { reader ->
            reader.asSequence().map { xclCell -> columns.columnMapper(xclCell, engine) }.forEach { writer.accept(it) }
        }
    }
}

fun <T> List<String>.columnMapper(cell: XCLCell<T>, engine: Interpreter): XCLCell<T> {
    engine.set("__cell__${cell.row}__${cell.col}", cell.value)
    return if (this.isEmpty()) {
        cell
    } else {
        val expr = this[cell.col].replace("$", "__cell__${cell.row}__")
        engine.eval("current = $expr")
        @Suppress("UNCHECKED_CAST")
        val value: T = engine.get("current") as T
        XCLCell(cell.row, cell.col, cell.type, value, cell.last)
    }
}
fun inExcelStream(input: File?, tab: String): CellProducer {
    return XSSFWorkbookReader(input?.inputStream() ?: System.`in`, tab)
}

fun inCSVStream(input: File?): CellProducer {
    return CSVReader(input?.inputStream() ?: System.`in`)
}

fun outCSVStream(output: File?): CellConsumer {
    return CSVWriter(output?.outputStream() ?: System.out)
}

fun outExcelStream(output: File?, tab: String): CellConsumer {
    return XSSFWorkbookWriter(output?.outputStream() ?: System.out, tab)
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
                XCLCell(row, col++, STRING, s)
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


class XSSFWorkbookReader(private val stream: InputStream, tab: String) : CellProducer {
    private val workbook = XSSFWorkbook(stream)
    private val sheet = workbook.getSheet(tab)
    override fun iterator(): Iterator<XCLCell<*>> {
        return sheet.asSequence().flatMap { row ->
            row.asSequence().map { cell ->
                val r = cell.rowIndex
                val c = cell.columnIndex
                val last = cell.columnIndex == cell.row.count() - 1
                when (cell.cellType) {
                    CellType.BLANK -> XCLCell(r, c, EMPTY, "", last)
                    CellType.NUMERIC -> if (DateUtil.isCellDateFormatted(cell)) {
                        XCLCell(r, c, DATE, cell.localDateTimeCellValue, last)
                    } else {
                        XCLCell(r, c, NUMBER, cell.numericCellValue, last)
                    }

                    CellType.STRING -> XCLCell(r, c, STRING, cell.stringCellValue, last)
                    CellType.FORMULA -> XCLCell(r, c, FORMULA, cell.cellFormula, last)
                    CellType.BOOLEAN -> XCLCell(r, c, BOOLEAN, cell.booleanCellValue, last)
                    CellType.ERROR -> XCLCell(r, c, ERROR, cell.errorCellValue, last)
                    CellType._NONE -> XCLCell(r, c, EMPTY, "", last)
                    else -> XCLCell(cell.rowIndex, c, EMPTY, "", last)
                }
            }
        }.iterator()
    }

    override fun close() {
        workbook.close()
        stream.close()
    }

}

class XSSFWorkbookWriter(private val stream: OutputStream, tab: String) : CellConsumer {
    private val workbook = SXSSFWorkbook()
    private val sheet = workbook.createSheet(tab)
    override fun accept(cell: XCLCell<*>) {
        val c = sheet.row(cell.row).cell(cell.col)
        when (cell.type) {
            NUMBER -> c.setCellValue(cell.value as Double)
            STRING -> c.setCellValue(cell.value as String)
            DATE -> c.setCellValue(cell.value as LocalDate)
            FORMULA -> c.setCellValue(cell.value as String)
            ERROR -> c.setCellValue(cell.value as String)
            BOOLEAN -> c.setCellValue(cell.value as Boolean)
            EMPTY -> c.setCellValue(cell.value as String)
        }
    }

    private fun SXSSFSheet.row(i: Int): SXSSFRow {
        return this.getRow(i) ?: this.createRow(i)
    }

    private fun SXSSFRow.cell(i: Int): SXSSFCell {
        return this.getCell(i, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK) ?: this.createCell(i)
    }

    override fun close() {
        try {
            workbook.write(stream);
            stream.close();
        } finally {
            // dispose of temporary files backing this workbook on disk
            workbook.dispose();
        }
    }

}