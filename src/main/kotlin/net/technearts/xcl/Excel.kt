package net.technearts.xcl

import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.ss.usermodel.CellType
import org.apache.poi.ss.usermodel.DateUtil
import org.apache.poi.ss.usermodel.Row
import org.apache.poi.xssf.streaming.SXSSFCell
import org.apache.poi.xssf.streaming.SXSSFRow
import org.apache.poi.xssf.streaming.SXSSFSheet
import org.apache.poi.xssf.streaming.SXSSFWorkbook
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.InputStream
import java.io.OutputStream
import java.time.LocalDate
import java.util.function.Predicate

class XSSFWorkbookWriter(
    private val stream: OutputStream,
    tab: String,
    private val start: Pair<Int, Int>?,
    private val end: Pair<Int, Int>?
) : CellConsumer {
    private val workbook = SXSSFWorkbook()
    private val sheet = workbook.createSheet(tab)
    override fun accept(cell: XCLCell<*>) {
        val c = sheet.row(cell.row).cell(cell.col)
        when (cell.type) {
            XCLCellType.NUMBER -> c.setCellValue(cell.value as Double)
            XCLCellType.STRING -> c.setCellValue(cell.value as String)
            XCLCellType.DATE -> c.setCellValue(cell.value as LocalDate)
            XCLCellType.FORMULA -> c.setCellValue(cell.value as String)
            XCLCellType.ERROR -> c.setCellValue(cell.value as String)
            XCLCellType.BOOLEAN -> c.setCellValue(cell.value as Boolean)
            XCLCellType.EMPTY -> c.setCellValue(cell.value as String)
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
            workbook.write(stream)
            stream.close()
        } finally {
            // dispose of temporary files backing this workbook on disk
            workbook.dispose()
        }
    }

}

class XSSFWorkbookReader(
    private val stream: InputStream,
    tab: String,
    private val start: Pair<Int, Int>?,
    private val end: Pair<Int, Int>?
) : CellProducer {
    private val workbook = XSSFWorkbook(stream)
    private val sheet = workbook.getSheet(tab)
    override fun iterator(): Iterator<XCLCell<*>> {
        return sheet.filter { row -> rowFilter(start, end).test(row) }
            .asSequence().flatMap { row ->
                row.filter { cell -> columnFilter(start, end).test(cell) }
                    .asSequence().map { cell ->
                        val r = cell.rowIndex
                        val c = cell.columnIndex
                        when (cell.cellType) {
                            CellType.BLANK -> XCLCell(r, c, XCLCellType.EMPTY, "")
                            CellType.NUMERIC -> if (DateUtil.isCellDateFormatted(cell)) {
                                XCLCell(r, c, XCLCellType.DATE, cell.localDateTimeCellValue)
                            } else {
                                XCLCell(r, c, XCLCellType.NUMBER, cell.numericCellValue)
                            }

                            CellType.STRING -> XCLCell(r, c, XCLCellType.STRING, cell.stringCellValue)
                            CellType.FORMULA -> XCLCell(r, c, XCLCellType.FORMULA, cell.cellFormula)
                            CellType.BOOLEAN -> XCLCell(r, c, XCLCellType.BOOLEAN, cell.booleanCellValue)
                            CellType.ERROR -> XCLCell(r, c, XCLCellType.ERROR, cell.errorCellValue)
                            CellType._NONE -> XCLCell(r, c, XCLCellType.EMPTY, "")
                            else -> XCLCell(cell.rowIndex, c, XCLCellType.EMPTY, "")
                        }
                    }
            }.iterator()
    }

    private fun rowFilter(start: Pair<Int, Int>?, end: Pair<Int, Int>?): Predicate<Row> =
        Predicate<Row> { row -> row.rowNum >= (start?.first ?: 0) && row.rowNum <= (end?.first ?: Int.MAX_VALUE) }

    private fun columnFilter(start: Pair<Int, Int>?, end: Pair<Int, Int>?): Predicate<Cell> =
        Predicate<Cell> { cell ->
            cell.columnIndex >= (start?.second ?: 0) && cell.columnIndex <= (end?.second ?: Int.MAX_VALUE)
        }

    override fun close() {
        workbook.close()
        stream.close()
    }

}