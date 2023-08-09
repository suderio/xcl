package net.technearts.xcl

import io.quarkus.arc.log.LoggerName
import io.quarkus.test.junit.QuarkusTest
import org.dhatim.fastexcel.VisibilityState
import org.dhatim.fastexcel.Workbook
import org.dhatim.fastexcel.reader.ReadableWorkbook
import org.jboss.logging.Logger
import org.junit.jupiter.api.Test
import java.io.*
import java.nio.file.Files
import java.nio.file.Paths
import java.util.stream.Stream
import org.dhatim.fastexcel.reader.CellType.*

fun getTestResources(fileName: String): File {
    return Paths.get("src", "test", "resources", fileName).toAbsolutePath().toFile()
}

@QuarkusTest
class ReplTest {


    @Throws(IOException::class)
    fun readExcel(fileLocation: File, tab: String): Map<Int, MutableList<String>> {
        val data: MutableMap<Int, MutableList<String>> = HashMap()
        FileInputStream(fileLocation).use { file ->
            ReadableWorkbook(file).use { wb ->
                val sheet = wb.findSheet(tab).get()
                sheet.openStream().use { rows ->
                    rows.forEach { r ->
                        data[r.rowNum] = ArrayList()
                        for (cell in r) {
                            data[r.rowNum]!!.add(cell.rawValue)
                        }
                    }
                }
            }
        }
        return data
    }


    @Test
    fun testReadRepl() {
        val file = getTestResources("read.xlsx")
        val tempFile = Files.createTempFile("temp", ".csv")
        println("Using ${file.absolutePath}")
        println("Writing $tempFile")
        repl(inExcelStream(file, "Planilha1", "", ""), outCSVStream(tempFile.toFile(), ",", "\r\n"), emptyList())
        val result = Files.readString(tempFile.toAbsolutePath())
        println(result)
        assert(result.startsWith("A,B,C"))
        assert(result.lines().filter { s -> s.isNotEmpty() }.size == 4)
    }

    @Test
    fun testWriteRepl() {
        val file = getTestResources("read.csv")
        val tempFile = Files.createTempFile("temp", ".xlsx")
        println("Using ${file.absolutePath}")
        println("Writing $tempFile")
        repl(inCSVStream(file,",", "\r\n"), outExcelStream(tempFile.toFile(), "PlanilhaA", "", ""), emptyList())
        val result = readExcel(tempFile.toFile(), "PlanilhaA")
        result.forEach{ (_, c) ->  println("$c") }
        result[0]?.let { assert(it.containsAll(listOf("A", "B", "C"))) }
        assert(result.size == 4)
    }

    @Test
    fun testReadReplWithTransformations() {
        val file = getTestResources("read.xlsx")
        val tempFile = Files.createTempFile("temp", ".csv")
        println("Using ${file.absolutePath}")
        println("Writing $tempFile")
        repl(inExcelStream(file, "Planilha1", "", ""), outCSVStream(tempFile.toFile(),",", "\r\n"), listOf("$0", "$1 + $1", "$2"))
        val result = Files.readString(tempFile.toAbsolutePath())
        println(result)
        assert(result.startsWith("A,BB,C"))
        assert(result.lines().filter { s -> s.isNotEmpty() }.size == 4)
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
                    when (cell.type) {
                        NUMBER -> XCLCell(r, c, XCLCellType.NUMBER, cell.asNumber())
                        STRING -> XCLCell(r, c, XCLCellType.STRING, cell.asString())
                        FORMULA -> XCLCell(r, c, XCLCellType.FORMULA, cell.formula)
                        ERROR -> XCLCell(r, c, XCLCellType.ERROR, cell.rawValue)
                        BOOLEAN -> XCLCell(r, c, XCLCellType.BOOLEAN, cell.asBoolean())
                        EMPTY -> XCLCell(r, c, XCLCellType.EMPTY, "")
                        else -> XCLCell(r, c, XCLCellType.ERROR, "")
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
        sheet.visibilityState = VisibilityState.VISIBLE
    }

    override fun accept(cell: XCLCell<*>) {
        // TODO tratar pelo tipo da cell
        sheet.value(cell.row, cell.col, cell.value.toString())
    }
}