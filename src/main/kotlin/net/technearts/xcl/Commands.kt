package net.technearts.xcl

import io.quarkus.arc.log.LoggerName
import io.quarkus.picocli.runtime.annotations.TopCommand
import org.jboss.logging.Logger
import picocli.CommandLine.*
import picocli.CommandLine.Help.Visibility.ALWAYS
import java.io.File
import java.util.regex.Pattern

@TopCommand
@Command(
    name = "sheeter",
    description = ["""
        A simple cli utility to read and write excel and csv files.
        Sheeter processes csv and xlsx files, allows for simple transformations,
        and works with pipes.
        
        Usage:

        # Takes an excel file and creates a csv file
        cat file.xlsx | xcl xl2csv >> result.csv
        
        # Takes a csv file with fields separated by a semicolon and pipes to stdout
        cat file.csv | xcl csv2xl --separator="\n" --delimiter=";" --sheet=tab02 >> result.xlsx
        
        # Takes an excel file and pipes to stdout only the first 3 columns
        cat file.xlsx | xcl xl2xl "$1, $2, $3" >> result.xlsx
        
        # Subcommands can be repeated:
        xcl xl2csv --in file.xlsx --out result.csv \
            csv2xl --in file.csv --out result.xlsx
        
        """],
    scope = ScopeType.INHERIT,
    mixinStandardHelpOptions = true,
    sortOptions = false,
    sortSynopsis = false,
    usageHelpAutoWidth = true,
    showAtFileInUsageHelp = true,
    exitCodeOnSuccess = 0,
    exitCodeOnInvalidInput = 1,
    subcommandsRepeatable = true,
    subcommands = [HelpCommand::class, CsvToExcelCommand::class, ExcelToCsvCommand::class, CsvToCsvCommand::class, ExcelToExcelCommand::class]
)
class MainCommand


@Command(
    name = "csv2xl", description = [
        """Creates a new excel file from the input. The input must
be a csv file, although its format is fully configurable.
The default configurations is a modified RFC 4180 that
works well with Excel csv files.
"""]
)
class CsvToExcelCommand : Runnable {
    @LoggerName("xcl")
    lateinit var log: Logger

    @Option(names = ["-i", "--in"], description = ["input file"])
    var inputFile: File? = null

    @Option(
        names = ["--inStartCell"],
        description = ["Starting cell in the RnCn format"],
        defaultValue = "R1C1",
        converter = [AddressConverter::class]
    )
    lateinit var inStartCell: Address

    @Option(
        names = ["--inEndCell"],
        description = ["End cell in the RnCn format"],
        defaultValue = "",
        converter = [AddressConverter::class]
    )
    lateinit var inEndCell: Address

    @Option(
        names = ["-d", "--delimiter"],
        description = ["CSV field delimiter"],
        defaultValue = ",",
        showDefaultValue = ALWAYS
    )
    lateinit var delimiter: String

    @Option(
        names = ["-s", "--separator"],
        description = ["CSV record separator"],
        defaultValue = "\r\n",
        showDefaultValue = ALWAYS
    )
    lateinit var separator: String

    @Option(names = ["-o", "--out"], description = ["output file"])
    var outputFile: File? = null

    @Option(
        names = ["--sheet"], description = ["Name of the sheet"], defaultValue = "Sheet1", showDefaultValue = ALWAYS
    )
    lateinit var sheet: String

    @Option(
        names = ["--outStartCell"],
        paramLabel = "Starting Cell",
        description = ["Starting cell"],
        defaultValue = "R1C1",
        showDefaultValue = ALWAYS,
        converter = [AddressConverter::class]
    )
    lateinit var outStartCell: Address

    @ArgGroup(exclusive = true)
    lateinit var headersOptions: HeadersOptions

    class HeadersOptions {
        @Option(
            names = ["--headers"],
            paramLabel = "New headers",
            description = ["The new headers in the out file"],
            defaultValue = "",
            showDefaultValue = ALWAYS
        )
        lateinit var headers: List<String>

        @Option(
            names = ["-removeHeaders"], paramLabel = "Remove headers",
            description = ["Remove the headers in the out file"],
            defaultValue = "false",
            showDefaultValue = ALWAYS
        )
        var removeHeaders = false
    }

    @Parameters(paramLabel = "columns", description = ["Your output columns"], arity = "0..*")
    var columns: List<String>? = null

    override fun run() {
        log.info("tab: $sheet")
        log.info("input file: ${inputFile ?: "stdin"}")
        log.info("output file: ${outputFile ?: "stdout"}")
        repl(
            inCSVStream(inputFile, delimiter, separator, inStartCell, inEndCell, headersOptions.removeHeaders),
            outExcelStream(outputFile, sheet, outStartCell, headersOptions.headers),
            columns ?: emptyList()
        )
    }

}

@Command(
    name = "xl2csv", description = [
        """Creates a new csv file from the input. The input must
be an excel file. The default csv format is a modified
RFC 4180 similar to Excel csv files.
"""]
)
class ExcelToCsvCommand : Runnable {

    @LoggerName("xcl")
    private lateinit var log: Logger

    @Option(names = ["-i", "--in"], description = ["input file"])
    var inputFile: File? = null

    @Option(names = ["--sheet"], description = ["Name of the sheet"], defaultValue = "Sheet1")
    lateinit var sheet: String

    @Option(
        names = ["--inStartCell"],
        description = ["Starting cell in the RnCn format"],
        defaultValue = "R1C1",
        converter = [AddressConverter::class]
    )
    lateinit var inStartCell: Address

    @Option(
        names = ["--inEndCell"],
        description = ["End cell in the RnCn format"],
        defaultValue = "",
        converter = [AddressConverter::class]
    )
    lateinit var inEndCell: Address

    @Option(names = ["-o", "--out"], description = ["output file"])
    var outputFile: File? = null

    @Option(names = ["-d", "--delimiter"], description = ["CSV field delimiter"], defaultValue = ",")
    lateinit var delimiter: String

    @Option(names = ["-s", "--separator"], description = ["CSV record separator"], defaultValue = "\r\n")
    lateinit var separator: String

    @ArgGroup(exclusive = true)
    lateinit var headersOptions: HeadersOptions

    class HeadersOptions {
        @Option(
            names = ["--headers"],
            paramLabel = "New headers",
            description = ["The new headers in the out file"],
            defaultValue = "",
            showDefaultValue = ALWAYS
        )
        lateinit var headers: List<String>

        @Option(
            names = ["-removeHeaders"], paramLabel = "Remove headers",
            description = ["Remove the headers in the out file"],
            defaultValue = "false",
            showDefaultValue = ALWAYS
        )
        var removeHeaders = false
    }

    @Parameters(paramLabel = "columns", description = ["Your output columns"], arity = "0..*")
    var columns: List<String>? = null

    override fun run() {
        log.info("Sheet: $sheet")
        log.info("input file: ${inputFile ?: "stdin"}")
        log.info("output file: ${outputFile ?: "stdout"}")
        repl(
            inExcelStream(inputFile, sheet, inStartCell, inEndCell, headersOptions.removeHeaders),
            outCSVStream(outputFile, delimiter, separator, headersOptions.headers),
            columns ?: emptyList()
        )
    }
}

@Command(
    name = "csv2csv",
    description = [
        """Creates a new csv file from the input. The input must
be a csv file, although the input and output formats are
fully configurable. The default configurations is a 
modified RFC 4180 that is similar to Excel csv files.
"""],
)
class CsvToCsvCommand : Runnable {

    @LoggerName("xcl")
    private lateinit var log: Logger

    @Option(names = ["-i", "--in"], description = ["input file"])
    var inputFile: File? = null

    @Option(
        names = ["--inStartCell"],
        description = ["Starting cell in the RnCn format"],
        defaultValue = "R1C1",
        converter = [AddressConverter::class]
    )
    lateinit var inStartCell: Address

    @Option(
        names = ["--inEndCell"],
        description = ["End cell in the RnCn format"],
        defaultValue = "",
        converter = [AddressConverter::class]
    )
    lateinit var inEndCell: Address

    @Option(names = ["-d", "--inDelimiter"], description = ["CSV field delimiter"], defaultValue = ",")
    lateinit var inDelimiter: String

    @Option(names = ["-s", "--inSeparator"], description = ["CSV record separator"], defaultValue = "\r\n")
    lateinit var inSeparator: String

    @Option(names = ["-o", "--out"], description = ["output file"])
    var outputFile: File? = null

    @Option(names = ["-e", "--outDelimiter"], description = ["CSV field delimiter"], defaultValue = ",")
    lateinit var outDelimiter: String

    @Option(names = ["-t", "--outSeparator"], description = ["CSV record separator"], defaultValue = "\r\n")
    lateinit var outSeparator: String

    @ArgGroup(exclusive = true)
    lateinit var headersOptions: HeadersOptions

    class HeadersOptions {
        @Option(
            names = ["--headers"],
            paramLabel = "New headers",
            description = ["The new headers in the out file"],
            defaultValue = "",
            showDefaultValue = ALWAYS
        )
        lateinit var headers: List<String>

        @Option(
            names = ["-removeHeaders"], paramLabel = "Remove headers",
            description = ["Remove the headers in the out file"],
            defaultValue = "false",
            showDefaultValue = ALWAYS
        )
        var removeHeaders = false
    }

    @Parameters(paramLabel = "columns", description = ["Your output columns"], arity = "0..*")
    var columns: List<String>? = null

    override fun run() {
        log.info("input file: ${inputFile ?: "stdin"}")
        log.info("output file: ${outputFile ?: "stdout"}")
        repl(
            inCSVStream(inputFile, inDelimiter, inSeparator, inStartCell, inEndCell, headersOptions.removeHeaders),
            outCSVStream(outputFile, outDelimiter, outSeparator, headersOptions.headers),
            columns ?: emptyList()
        )
    }
}

@Command(
    name = "xl2xl",
    description = [
        """Creates a new excel file from the input. The input must
be another excel file.
"""],
)
class ExcelToExcelCommand : Runnable {

    @LoggerName("xcl")
    private lateinit var log: Logger

    @Option(names = ["-i", "--in"], description = ["input file"])
    var inputFile: File? = null

    @Option(names = ["--sheet"], description = ["Name of the sheet"], defaultValue = "Sheet1")
    lateinit var sheet: String

    @Option(
        names = ["--inStartCell"],
        description = ["Starting cell in the RnCn format"],
        defaultValue = "R1C1",
        converter = [AddressConverter::class]
    )
    lateinit var inStartCell: Address

    @Option(
        names = ["--inEndCell"],
        description = ["End cell in the RnCn format"],
        defaultValue = "",
        converter = [AddressConverter::class]
    )
    lateinit var inEndCell: Address

    @Option(names = ["-o", "--out"], description = ["output file"])
    var outputFile: File? = null

    @Option(
        names = ["--outStartCell"],
        description = ["Starting cell in the RnCn format"],
        defaultValue = "R1C1",
        converter = [AddressConverter::class]
    )
    lateinit var outStartCell: Address

    @ArgGroup(exclusive = true)
    lateinit var headersOptions: HeadersOptions

    class HeadersOptions {
        @Option(
            names = ["--headers"],
            paramLabel = "New headers",
            description = ["The new headers in the out file"],
            defaultValue = "",
            showDefaultValue = ALWAYS
        )
        lateinit var headers: List<String>

        @Option(
            names = ["-removeHeaders"], paramLabel = "Remove headers",
            description = ["Remove the headers in the out file"],
            defaultValue = "false",
            showDefaultValue = ALWAYS
        )
        var removeHeaders = false
    }

    @Parameters(paramLabel = "columns", description = ["Your output columns"], arity = "0..*")
    var columns: List<String>? = null

    override fun run() {
        log.info("Sheet: $sheet")
        log.info("input file: ${inputFile ?: "stdin"}")
        log.info("output file: ${outputFile ?: "stdout"}")
        repl(
            inExcelStream(inputFile, sheet, inStartCell, inEndCell, headersOptions.removeHeaders),
            outExcelStream(outputFile, sheet, outStartCell, headersOptions.headers),
            columns ?: emptyList()
        )
    }
}

class AddressConverter : ITypeConverter<Address?> {
    private val r1c1Pattern: Pattern = Pattern.compile("R\\d+C\\d+", Pattern.CASE_INSENSITIVE)
    override fun convert(address: String?): Address? {
        if (address?.isNotBlank() == true && r1c1Pattern.matcher(address).matches()) {
            val result = address.split('r', 'c', ignoreCase = true, limit = 0)
                .map(String::toInt)
                .filterIndexed { i, _ -> i <= 1 }
            //.map{i -> i - 1}
            return Address(result[0], result[1])
        }
        return null
    }

}