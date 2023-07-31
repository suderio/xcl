package net.technearts.xcl

import io.quarkus.test.junit.QuarkusTest
import org.junit.jupiter.api.Test
import java.io.File
import java.nio.file.Files
import java.nio.file.Paths

fun getTestResources(fileName: String): File {
    return Paths.get("src", "test", "resources", fileName).toAbsolutePath().toFile()
}

@QuarkusTest
class ReplTest {

    @Test
    fun testReadRepl() {
        val file = getTestResources("read.xlsx")
        val tempFile = Files.createTempFile("temp", "txt")
        println("Using ${file.absolutePath}")
        println("Writing $tempFile")
        repl(inStream(file, "Planilha1"), outStream(tempFile.toFile()))
        val result = Files.readString(tempFile.toAbsolutePath())
        print(result)
        assert(result.startsWith("A,B,C"))
    }
}