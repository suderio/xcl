package net.technearts.xcl

import io.quarkus.arc.log.LoggerName
import io.quarkus.picocli.runtime.annotations.TopCommand
import org.jboss.logging.Logger
import picocli.CommandLine.*
import java.io.File

/**
 * cat file.xlsx | xcl
 * outputs the content of the first tab as csv
 */

@TopCommand
@Command(
    name = "xcl",
    mixinStandardHelpOptions = true,
    description = ["Reads one sheet of an excel file and outputs as a CSV.", "This is the default command."],
    subcommands = [CreateCommand::class, UpdateCommand::class]
)
class ReadCommand : Runnable {

    @LoggerName("xcl")
    private lateinit var log: Logger

    @Option(names = ["-t", "--tab"], description = ["Name of the tab"], defaultValue = "tab1")
    lateinit var tab: String

    @Option(names = ["-i", "--in"], description = ["input file"])
    var inputFile: File? = null

    @Option(names = ["-o", "--out"], description = ["output file"])
    var outputFile: File? = null

    @Parameters(paramLabel = "columns", description = ["Your output columns"], arity = "0..*")
    var columns: List<String>? = null

    override fun run() {
        log.info("tab: $tab")
        log.info("input file: ${inputFile ?: "stdin"}")
        log.info("output file: ${outputFile ?: "stdout"}")
        repl(inExcelStream(inputFile, tab), outCSVStream(outputFile), columns?: emptyList())
    }
}

/**
 *
 * cat file.csv | xcl create --template template.xsl --line-separator="\n" \
 * --field-separator="," --cell=B1 --tab=tab02 >> result.xls
 */
@Command(name = "create", mixinStandardHelpOptions = true, description = ["Creates a new excel file from the input"])
class CreateCommand : Runnable {
    @LoggerName("xcl")
    lateinit var log: Logger

    @Option(names = ["-t", "--tab"], description = ["Name of the tab"], defaultValue = "tab1")
    lateinit var tab: String

    @Option(names = ["-i", "--in"], description = ["input file"])
    var inputFile: File? = null

    @Option(names = ["-o", "--out"], description = ["output file"])
    var outputFile: File? = null

    @Parameters(paramLabel = "columns", description = ["Your output columns"], arity = "0..*")
    var columns: List<String>? = null

    override fun run() {
        log.info("tab: $tab")
        log.info("input file: ${inputFile ?: "stdin"}")
        log.info("output file: ${outputFile ?: "stdout"}")
        repl(inCSVStream(inputFile), outExcelStream(outputFile, tab), columns?: emptyList())
    }

}

/**
 *
 * cat file.csv | xcl create --template template.xsl --line-separator="\n" \
 * --field-separator="," --cell=B1 --tab=tab02 >> result.xls
 */
@Command(name = "update", mixinStandardHelpOptions = true)
class UpdateCommand : Runnable {
    @LoggerName("xcl")
    lateinit var log: Logger

    @Option(names = ["-t", "--tab"], description = ["Name of the tab"], defaultValue = "tab1")
    lateinit var tab: String

    @Option(names = ["-c", "--cell"], description = ["Starting cell"], defaultValue = "A1")
    lateinit var cell: String

    @Option(names = ["-i", "--in"], description = ["input file"])
    var inputFile: File? = null

    @Option(names = ["-o", "--out"], description = ["output file"])
    var outputFile: File? = null

    @Option(
        names = ["--head"],
        description = ["Reads only the first lines of input", "or (if negative) skips the initial lines"],
        defaultValue = "0"
    )
    var head: Int = 0

    @Option(
        names = ["--tail"],
        description = ["Reads only the last lines of input", "or (if negative) skips the final lines"],
        defaultValue = "0"
    )
    var tail: Int = 0

    @Parameters(paramLabel = "columns", description = ["Your output columns"], arity = "0..*")
    var columns: List<String>? = null

    override fun run() {
        log.info("tab: $tab")
        log.info("cell: $cell")
        log.info("input file: ${inputFile ?: "stdin"}")
        log.info("output file: ${outputFile ?: "stdout"}")
        TODO("Not yet implemented")
    }
}