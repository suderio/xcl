package net.technearts.xcl

import io.quarkus.test.junit.QuarkusTest
import org.dhatim.fastexcel.reader.ReadableWorkbook
import org.junit.jupiter.api.Test
import java.io.File
import java.io.FileInputStream
import java.io.IOException
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
        val tempFile = Files.createTempFile("temp", ".txt")
        println("Using ${file.absolutePath}")
        println("Writing $tempFile")
        repl(inExcelStream(file, "Planilha1"), outCSVStream(tempFile.toFile()))
        val result = Files.readString(tempFile.toAbsolutePath())
        println(result)
        assert(result.startsWith("A,B,C"))
    }

    @Test
    fun testWriteRepl() {
        val file = getTestResources("read.csv")
        val tempFile = Files.createTempFile("temp", ".xlsx")
        println("Using ${file.absolutePath}")
        println("Writing $tempFile")
        repl(inCSVStream(file), outExcelStream(tempFile.toFile(), "PlanilhaA"))
        val result = readExcel(tempFile.toFile(), "PlanilhaA")
        result.forEach{ (_, c) ->  println("$c") }
        result[0]?.let { assert(it.containsAll(listOf("A", "B", "C"))) }
    }
}