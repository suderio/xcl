package net.technearts.xcl

import org.apache.commons.csv.CSVFormat
import org.apache.commons.csv.CSVPrinter
import org.apache.commons.csv.CSVRecord
import org.apache.commons.csv.DuplicateHeaderMode
import java.io.*

class CSVWriter(stream: OutputStream, format: CSVFormat) : FilterWriter(BufferedWriter(OutputStreamWriter(stream))),
    CellConsumer {
    private val printer = CSVPrinter(this, format)
    private var currentRow: Int? = null
    private var record = mutableListOf<String>()

    override fun accept(cell: XCLCell<*>) {
        if ((currentRow ?: cell.row) == cell.row) {
            record += cell.value.toString()
        } else {
            printer.printRecord(record)
            record.clear()
        }
        currentRow = cell.row
    }

    override fun close() {
        // This is probably smelly as hell
        printer.printRecord(record)
        super.close()
    }
}

class CSVReader(stream: InputStream, private val format: CSVFormat) : BufferedReader(InputStreamReader(stream)),
    CellProducer {
    override fun iterator(): Iterator<XCLCell<*>> {
        return format.parse(this)
            .asSequence()
            .flatMapIndexed { row: Int, record: CSVRecord ->
                var col = 0
                record.asSequence().map { field -> XCLCell(row, col++, XCLCellType.STRING, field) }
            }.iterator()
    }
}

fun builder(
    delimiter: String = ",",
    separator: String = "\r\n",
    commentMarker: Char? = null,
    escape: Char? = null,
    quote: Char? = null
): CSVFormat.Builder {
    return CSVFormat.Builder.create()
        .setDelimiter(delimiter)
        .setCommentMarker(commentMarker)
        .setEscape(escape)
        .setQuote(quote)
        .setRecordSeparator(separator)
        .setTrim(true)
        .setQuoteMode(null)
        .setNullString(null)
        .setIgnoreSurroundingSpaces(true)
        .setIgnoreEmptyLines(true)
        .setDuplicateHeaderMode(DuplicateHeaderMode.ALLOW_ALL)
        .setAllowMissingColumnNames(true)
        .setAutoFlush(true)
}