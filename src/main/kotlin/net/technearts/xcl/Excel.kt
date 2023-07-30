package net.technearts.xcl

import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.ss.usermodel.CellType.*
import org.apache.poi.ss.usermodel.DateUtil
import org.apache.poi.ss.usermodel.Row
import org.apache.poi.xssf.usermodel.XSSFSheet


class CellIterator(sheet: XSSFSheet) : Iterator<Cell> {
    private val rows = sheet.iterator()
    private lateinit var row: Row
    private lateinit var cells: MutableIterator<Cell>

    init {
        if (rows.hasNext()) {
            row = rows.next()
            cells = row.cellIterator()
        }
    }

    override fun hasNext(): Boolean {
        return rows.hasNext() || cells.hasNext()
    }

    override fun next(): Cell {
        if (cells.hasNext()) {
            return cells.next()
        } else {
            while (rows.hasNext()) {
                row = rows.next()
                cells = row.cellIterator()
                return next()
            }
        }
        throw NoSuchElementException()
    }

    fun endOfLine() = !cells.hasNext()
}

fun Cell.getCellValue() = when (this.cellType) {
    STRING -> this.richStringCellValue.toString()
    NUMERIC -> if (DateUtil.isCellDateFormatted(this)) {
        this.dateCellValue.toString()
    } else {
        this.numericCellValue.toString()
    }

    BOOLEAN -> this.booleanCellValue.toString()
    FORMULA -> this.cellFormula.toString()
    ERROR -> this.cellType.name
    BLANK, _NONE -> ""
    else -> ""
}