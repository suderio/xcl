package net.technearts.xcl

import io.quarkus.arc.log.LoggerName
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import org.jboss.logging.Logger
import java.io.*

const val FS = '\u00F1'
const val LF = '\u000A'

fun inWorkbookStream(input: File?, tab: String): Reader = WorkbookReader(input?.inputStream() ?: System.`in`, tab)

fun outCSVStream(output: File?): Writer = CSVWriter(output?.outputStream() ?: System.out)

fun inCSVStream(input: File?): Reader = CSVReader(input?.inputStream() ?: System.`in`)

fun outWorkbookStream(output: File?, tab: String): Writer = WorkbookWriter(output?.outputStream() ?: System.out, tab)
fun repl(input: Reader, output: Writer) {
    BufferedWriter(output).use { writer ->
        BufferedReader(input).use { reader ->
            reader.lines().forEach {
                writer.append(it)
                writer.newLine()
            }
        }
    }
}

class WorkbookReader(
    stream: InputStream, tab: String
) : FilterReader(InputStreamReader(stream)) {
    private val workbook = XSSFWorkbook(stream)
    private val sheet = workbook.getSheet(tab)
    private val iterator = CellIterator(sheet)

    @LoggerName("xcl")
    lateinit var log: Logger
    override fun read(p0: CharArray, p1: Int, p2: Int): Int {
        return if (iterator.hasNext()) {
            val s = "${iterator.next().getCellValue()}${if (iterator.endOfLine()) LF else FS}"
            if (s.length > p2)
                log.error("Error in the WorkbookReader buffer.\nReading ($s)\nMax ($p2)")
            s.reader().read(p0, p1, p2)
        } else {
            -1
        }
    }

    override fun close() {
        workbook.close()
    }

}

class WorkbookWriter(
    private val stream: OutputStream,
    tab: String,
    private val fieldSeparator: Char = ',',
    private val lineSeparator: Char = '\n'
) : FilterWriter(OutputStreamWriter(stream)) {
    private val workbook = XSSFWorkbook()
    private val sheet = workbook.createSheet(tab)

    override fun close() {
        workbook.close()
        super.close()
    }

    override fun flush() {
        workbook.write(stream)
        super.flush()
    }

    override fun write(p0: CharArray, p1: Int, p2: Int) {
        TODO()
    }

}

class CSVReader(
    stream: InputStream,
    private val fieldSeparator: Char = ',',
    private val lineSeparator: Char = '\n'
) : FilterReader(InputStreamReader(stream)) {
    override fun read(p0: CharArray, p1: Int, p2: Int): Int {
        return this.`in`.read(separator(p0), p1, p2)
    }

    private fun separator(p0: CharArray) =
        p0.map {
            when (it) {
                fieldSeparator -> FS
                lineSeparator -> LF
                else -> it
            }
        }.toCharArray()
}

class CSVWriter(
    stream: OutputStream,
    private val fieldSeparator: Char = ',',
    private val lineSeparator: Char = '\n'
) : FilterWriter(OutputStreamWriter(stream)) {
    override fun write(p0: CharArray, p1: Int, p2: Int) {
        this.out.write(separators(p0), p1, p2)
    }

    private fun separators(p0: CharArray) =
        p0.map {
            when (it) {
                FS -> fieldSeparator
                LF -> lineSeparator
                else -> it
            }
        }.toCharArray()
}

