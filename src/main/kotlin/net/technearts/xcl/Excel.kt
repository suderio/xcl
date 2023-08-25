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
import java.time.LocalDateTime
import java.util.function.Predicate

class XSSFWorkbookWriter(
    private val stream: OutputStream,
    tab: String,
    private val start: Address,
    private val headers: List<String>
) : CellConsumer {
    private val workbook = SXSSFWorkbook()
    private val sheet = workbook.createSheet(tab)
    private val dateCellStyle = workbook.createCellStyle()

    init {
        dateCellStyle.dataFormat = workbook.creationHelper.createDataFormat().getFormat("dd/mm/yyyy")
    }

    override fun accept(cell: XCLCell<*>) {
        val xCell = sheet.row(cell.address.row + start.row).cell(cell.address.column + start.column)
        if (headers.isNotEmpty() && cell.address.row == 0) {
            xCell.setCellValue(headers[cell.address.column])
            xCell.cellType = CellType.STRING
            return
        }
        when (cell.type) {
            XCLCellType.NUMBER -> {
                xCell.setCellValue(cell.value as Double)
                xCell.cellType = CellType.NUMERIC
            }

            XCLCellType.STRING -> {
                xCell.setCellValue(cell.value as String)
                xCell.cellType = CellType.STRING
            }

            XCLCellType.DATE -> {
                xCell.setCellValue(cell.value as LocalDateTime)
                xCell.cellType = CellType.NUMERIC
                xCell.cellStyle = dateCellStyle
            }

            XCLCellType.FORMULA -> {
                xCell.setCellValue(cell.value as String)
                xCell.cellType = CellType.FORMULA
            }

            XCLCellType.ERROR -> {
                xCell.setCellValue(cell.value as String)
                xCell.cellType = CellType.ERROR
            }

            XCLCellType.BOOLEAN -> {
                xCell.setCellValue(cell.value as Boolean)
                xCell.cellType = CellType.BOOLEAN
            }

            XCLCellType.EMPTY -> {
                xCell.setCellValue(cell.value as String)
                xCell.cellType = CellType.BLANK
            }
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
    private val start: Address?,
    private val end: Address?,
    private val removeHeaders: Boolean
) : CellProducer {
    private val workbook = XSSFWorkbook(stream)
    private val sheet = workbook.getSheet(tab)
    override fun iterator(): Iterator<XCLCell<*>> {
        return sheet.filter { row -> rowFilter(start, end).test(row) }
            .asSequence().flatMap { row ->
                row.filter { cell -> columnFilter(start, end).test(cell) }
                    .asSequence().map { cell ->
                        val address = Address(cell.rowIndex, cell.columnIndex)
                        when (cell.cellType) {
                            CellType.BLANK -> XCLCell(address, XCLCellType.EMPTY, "")
                            CellType.NUMERIC -> if (DateUtil.isCellDateFormatted(cell)) {
                                XCLCell(address, XCLCellType.DATE, cell.localDateTimeCellValue)
                            } else {
                                XCLCell(address, XCLCellType.NUMBER, cell.numericCellValue)
                            }

                            CellType.STRING -> XCLCell(address, XCLCellType.STRING, cell.stringCellValue)
                            CellType.FORMULA -> XCLCell(address, XCLCellType.FORMULA, cell.cellFormula)
                            CellType.BOOLEAN -> XCLCell(address, XCLCellType.BOOLEAN, cell.booleanCellValue)
                            CellType.ERROR -> XCLCell(address, XCLCellType.ERROR, cell.errorCellValue)
                            CellType._NONE -> XCLCell(address, XCLCellType.EMPTY, "")
                            else -> XCLCell(address, XCLCellType.EMPTY, "")
                        }
                    }
            }.iterator()
    }

    private fun rowFilter(start: Address?, end: Address?): Predicate<Row> {
        return Predicate<Row> { row ->
            (row.rowNum >= (start?.row ?: 0)
                    && row.rowNum <= (end?.row ?: Int.MAX_VALUE)
                    && (!removeHeaders || row.rowNum != (start?.row ?: 0)))
        }
    }

    private fun columnFilter(start: Address?, end: Address?): Predicate<Cell> =
        Predicate<Cell> { cell ->
            cell.columnIndex >= (start?.column ?: 0) && cell.columnIndex <= (end?.column ?: Int.MAX_VALUE)
        }

    override fun close() {
        workbook.close()
        stream.close()
    }

}