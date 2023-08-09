package net.technearts.xcl

import io.quarkus.arc.log.LoggerName
import io.quarkus.picocli.runtime.annotations.TopCommand
import org.jboss.logging.Logger
import picocli.CommandLine.*
import java.io.File
import kotlin.system.exitProcess

/**
 * cat file.xlsx | xcl
 * outputs the content of the first tab as csv
 */

@TopCommand
@Command(
    name = "xcl",
    mixinStandardHelpOptions = true,
    helpCommand = true,
    description = ["Reads one sheet of an excel file and outputs as a CSV.", "This is the default command."],
    subcommands = [CsvToExcelCommand::class, ExcelToCsvCommand::class, CsvToCsvCommand::class, ExcelToExcelCommand::class]
)
class HelpCommand : Runnable {

    override fun run() {
        exitProcess(0)
    }
}

/**
 *
 * cat file.csv | xcl create --template template.xsl --line-separator="\n" \
 * --field-separator="," --cell=B1 --tab=tab02 >> result.xls
 */
@Command(name = "create", mixinStandardHelpOptions = true, description = ["Creates a new excel file from the input"])
class CsvToExcelCommand : Runnable {
    @LoggerName("xcl")
    lateinit var log: Logger

    @Option(names = ["--sheet"], description = ["Name of the sheet"], defaultValue = "Sheet1")
    lateinit var sheet: String

    @Option(names = ["-d", "--delimiter"], description = ["CSV field delimiter"], defaultValue = ",")
    lateinit var delimiter: String

    @Option(names = ["-s", "--separator"], description = ["CSV record separator"], defaultValue = "\r\n")
    lateinit var separator: String

    @Option(names = ["-i", "--in"], description = ["input file"])
    var inputFile: File? = null

    @Option(names = ["-o", "--out"], description = ["output file"])
    var outputFile: File? = null

    @Option(names = ["--outStartCell"], description = ["Starting cell"], defaultValue = "R1C1")
    lateinit var outStartCell: String

    @Option(names = ["--outEndCell"], description = ["Starting cell"], defaultValue = "")
    lateinit var outEndCell: String

    @Parameters(paramLabel = "columns", description = ["Your output columns"], arity = "0..*")
    var columns: List<String>? = null

    override fun run() {
        log.info("tab: $sheet")
        log.info("input file: ${inputFile ?: "stdin"}")
        log.info("output file: ${outputFile ?: "stdout"}")
        repl(
            inCSVStream(inputFile, delimiter, separator),
            outExcelStream(outputFile, sheet, outStartCell, outEndCell),
            columns ?: emptyList()
        )
    }

}

/**
 *
 * cat file.xlsx | xcl xl2csv >> result.csv
 */
@Command(name = "xl2csv", mixinStandardHelpOptions = true)
class ExcelToCsvCommand : Runnable {

    @LoggerName("xcl")
    private lateinit var log: Logger

    @Option(names = ["-d", "--delimiter"], description = ["CSV field delimiter"], defaultValue = ",")
    lateinit var delimiter: String

    @Option(names = ["-s", "--separator"], description = ["CSV record separator"], defaultValue = "\r\n")
    lateinit var separator: String

    @Option(names = ["--sheet"], description = ["Name of the sheet"], defaultValue = "Sheet1")
    lateinit var sheet: String

    @Option(names = ["-i", "--in"], description = ["input file"])
    var inputFile: File? = null

    @Option(names = ["-o", "--out"], description = ["output file"])
    var outputFile: File? = null

    @Option(names = ["--inStartCell"], description = ["Starting cell in the RnCn format"], defaultValue = "R1C1")
    lateinit var inStartCell: String

    @Option(names = ["--inEndCell"], description = ["End cell in the RnCn format"], defaultValue = "")
    lateinit var inEndCell: String

    @Parameters(paramLabel = "columns", description = ["Your output columns"], arity = "0..*")
    var columns: List<String>? = null

    override fun run() {
        log.info("Sheet: $sheet")
        log.info("input file: ${inputFile ?: "stdin"}")
        log.info("output file: ${outputFile ?: "stdout"}")
        repl(
            inExcelStream(inputFile, sheet, inStartCell, inEndCell),
            outCSVStream(outputFile, delimiter, separator),
            columns ?: emptyList()
        )
    }
}

@Command(
    name = "csv2csv",
    description = ["Reads a csv file and writes to another."],
)
class CsvToCsvCommand : Runnable {

    @LoggerName("xcl")
    private lateinit var log: Logger

    @Option(names = ["-d", "--inDelimiter"], description = ["CSV field delimiter"], defaultValue = ",")
    lateinit var inDelimiter: String

    @Option(names = ["-s","--inSeparator"], description = ["CSV record separator"], defaultValue = "\r\n")
    lateinit var inSeparator: String

    @Option(names = ["-e", "--outDelimiter"], description = ["CSV field delimiter"], defaultValue = ",")
    lateinit var outDelimiter: String

    @Option(names = ["-t", "--outSeparator"], description = ["CSV record separator"], defaultValue = "\r\n")
    lateinit var outSeparator: String

    @Option(names = ["-i", "--in"], description = ["input file"])
    var inputFile: File? = null

    @Option(names = ["-o", "--out"], description = ["output file"])
    var outputFile: File? = null

    @Parameters(paramLabel = "columns", description = ["Your output columns"], arity = "0..*")
    var columns: List<String>? = null

    override fun run() {
        log.info("input file: ${inputFile ?: "stdin"}")
        log.info("output file: ${outputFile ?: "stdout"}")
        repl(
            inCSVStream(inputFile, inDelimiter, inSeparator),
            outCSVStream(outputFile, outDelimiter, outSeparator),
            columns ?: emptyList()
        )
    }
}

@Command(
    name = "xl2xl",
    description = ["Reads one sheet of an excel file and outputs as another excel file."],
)
class ExcelToExcelCommand : Runnable {

    @LoggerName("xcl")
    private lateinit var log: Logger

    @Option(names = ["--sheet"], description = ["Name of the sheet"], defaultValue = "Sheet1")
    lateinit var sheet: String

    @Option(names = ["-i", "--in"], description = ["input file"])
    var inputFile: File? = null

    @Option(names = ["-o", "--out"], description = ["output file"])
    var outputFile: File? = null

    @Option(names = ["--inStartCell"], description = ["Starting cell in the RnCn format"], defaultValue = "R1C1")
    lateinit var inStartCell: String

    @Option(names = ["--inEndCell"], description = ["End cell in the RnCn format"], defaultValue = "")
    lateinit var inEndCell: String

    @Option(names = ["--outStartCell"], description = ["Starting cell in the RnCn format"], defaultValue = "R1C1")
    lateinit var outStartCell: String

    @Option(names = ["--outEndCell"], description = ["End cell in the RnCn format"], defaultValue = "")
    lateinit var outEndCell: String

    @Parameters(paramLabel = "columns", description = ["Your output columns"], arity = "0..*")
    var columns: List<String>? = null

    override fun run() {
        log.info("Sheet: $sheet")
        log.info("input file: ${inputFile ?: "stdin"}")
        log.info("output file: ${outputFile ?: "stdout"}")
        repl(
            inExcelStream(inputFile, sheet, inStartCell, inEndCell),
            outExcelStream(outputFile, sheet, outStartCell, outEndCell),
            columns ?: emptyList()
        )
    }
}

