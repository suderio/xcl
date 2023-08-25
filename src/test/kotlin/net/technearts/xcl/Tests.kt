package net.technearts.xcl

import io.quarkus.test.junit.QuarkusTest
import org.dhatim.fastexcel.reader.CellType.*
import org.dhatim.fastexcel.reader.ReadableWorkbook
import org.junit.jupiter.api.Test
import java.io.*
import java.nio.file.Files
import java.nio.file.Paths

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
        println("xlsx -> csv")
        println("Using ${file.absolutePath}")
        println("Writing $tempFile")
        repl(
            inExcelStream(file, "Planilha1", Address(0, 0), Address(3, 3)),
            outCSVStream(tempFile.toFile(), ",", "\r\n"),
            emptyList(), true
        )
        val result = Files.readString(tempFile.toAbsolutePath())
        println(result)
        assert(result.startsWith("A,B,C"))
        assert(result.lines().filter { s -> s.isNotEmpty() }.size == 4)
        assert(result.lines().last{ line -> line.isNotEmpty() }.startsWith("teste3,3.3"))
    }

    @Test
    fun testWriteRepl() {
        val file = getTestResources("read.csv")
        val tempFile = Files.createTempFile("temp", ".xlsx")
        println("csv -> xlsx")
        println("Using ${file.absolutePath}")
        println("Writing $tempFile")
        repl(
            inCSVStream(file, ",", "\r\n", Address(0, 0), Address(3, 3)),
            outExcelStream(tempFile.toFile(), "PlanilhaA"),
            emptyList(), true
        )
        val result = readExcel(tempFile.toFile(), "PlanilhaA")
        result.forEach { (_, c) -> println("$c") }
        assert(result[1]!!.containsAll(listOf("A", "B", "C")))
        assert(result.size == 4)
        assert(result[4]!!.contains("teste3"))
    }

    @Test
    fun testWriteReplChangingHeaders() {
        val file = getTestResources("read.xlsx")
        val tempFile = Files.createTempFile("temp", ".xlsx")
        println("xlsx -> xlsx")
        println("Using ${file.absolutePath}")
        println("Writing $tempFile")
        repl(
            inExcelStream(file, "Planilha1", Address(0, 0), Address(3, 3)),
            outExcelStream(tempFile.toFile(), "PlanilhaA", headers = listOf("AAA", "BBB", "CCC")),
            emptyList(), true
        )
        val result = readExcel(tempFile.toFile(), "PlanilhaA")
        result.forEach { (_, c) -> println("$c") }
        assert(result[1]!!.containsAll(listOf("AAA", "BBB", "CCC")))
        assert(result.size == 4)
        assert(result[4]!!.contains("teste3"))
    }

    @Test
    fun testReadReplWithTransformations() {
        val file = getTestResources("read.xlsx")
        val tempFile = Files.createTempFile("temp", ".csv")
        println("xlsx -> csv (\"\$0\", \"\$1 + \$1\", \"\$2\")")
        println("Using ${file.absolutePath}")
        println("Writing $tempFile")
        repl(
            inExcelStream(file, "Planilha1", Address(0, 0), Address(3, 3)),
            outCSVStream(tempFile.toFile(), ",", "\r\n"),
            listOf("$0", "$1 + $1", "$2"),
            true
        )
        val result = Files.readString(tempFile.toAbsolutePath())
        println(result)
        assert(result.startsWith("A,BB,C"))
        assert(result.lines().filter { s -> s.isNotEmpty() }.size == 4)
        assert(result.lines().last { line -> line.isNotEmpty() }.startsWith("teste3,6.6"))
    }

    @Test
    fun testReadReplWithTransformationsAddingColumn() {
        val file = getTestResources("read.xlsx")
        val tempFile = Files.createTempFile("temp", ".csv")
        println("xlsx -> csv \"\$0\", \"\$1\", \"\$2\", \"\$1 + \$1\"")
        println("Using ${file.absolutePath}")
        println("Writing $tempFile")
        repl(
            inExcelStream(file, "Planilha1", Address(0, 0), Address(3, 3)),
            outCSVStream(tempFile.toFile(), ",", "\r\n"),
            listOf("$0", "$1", "$2", "$1 + $1"),
            true
        )
        val result = Files.readString(tempFile.toAbsolutePath())
        println(result)
        assert(result.startsWith("A,B,C,BB"))
        assert(result.lines().filter { s -> s.isNotEmpty() }.size == 4)
        assert(result.lines().last{ line -> line.isNotEmpty() }.startsWith("teste3,3.3"))
        assert(result.lines().last{ line -> line.isNotEmpty() }.endsWith(",6.6"))
    }
}
