package net.technearts.xcl

import io.quarkus.test.junit.QuarkusTest
import org.junit.jupiter.api.Test
import java.io.File
import java.nio.file.Paths

fun getTestResources(fileName: String): File {
    return Paths.get("src", "test", "resources", fileName).toAbsolutePath().toFile()
}

@QuarkusTest
class ReplTest {

    @Test
    fun testRepl() {
        println(getTestResources("read.xlsx").absolutePath)
    }
}