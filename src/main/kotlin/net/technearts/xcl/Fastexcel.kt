package net.technearts.xcl

import org.dhatim.fastexcel.Workbook
import org.dhatim.fastexcel.reader.Cell
import org.dhatim.fastexcel.reader.CellType.*
import org.dhatim.fastexcel.reader.ReadableWorkbook
import java.io.InputStream
import java.io.OutputStream
import java.time.LocalDate


fun read(input: InputStream, tab: String) {
    ReadableWorkbook(input).use { wb ->
        val sheet = wb.findSheet(tab).get()
        sheet.openStream().use { rows ->
            rows.forEach { r ->
                val num = r.getCellAsNumber(0).orElse(null);
                val str = r.getCellAsString(1).orElse(null);
                val date = r.getCellAsDate(2).orElse(null);
            }
        }
    }
}

fun write(output: OutputStream, tab: String) {
    val wb = Workbook(output, "xcl", "1.0")
    val ws = wb.newWorksheet(tab)
    ws.value(0, 0, 0)
    ws.value(0, 1, "1")
    ws.value(0, 2, LocalDate.now())
}

fun Cell.value(): Any {
    return when (this.type) {
        BOOLEAN -> this.asBoolean()
        NUMBER -> this.asNumber()
        STRING -> this.asString()
        FORMULA -> this.formula
        ERROR -> this.value
        EMPTY -> ""
        else -> ""
    }
    TODO("Tratar Date")
}