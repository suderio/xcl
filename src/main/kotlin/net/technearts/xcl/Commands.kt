package net.technearts.xcl

import io.quarkus.arc.log.LoggerName
import io.quarkus.picocli.runtime.annotations.TopCommand
import picocli.CommandLine.*
import org.jboss.logging.Logger
import java.io.File

/**
 * xcl << file.xlsx
 * outputs the content of the first tab as csv
 */
@TopCommand
@Command(name = "xcl", mixinStandardHelpOptions = true, subcommands = [Create::class, Update::class])
class Read : Runnable {

    @LoggerName("xcl")
    lateinit var log: Logger

    @Option(names = ["-t", "--tab"], description = ["Name of the tab"], defaultValue = "tab1")
    var tab: String = "tab1"

    @Option(names = ["-c", "--cell"], description = ["Starting cell"], defaultValue = "A1")
    var cell: String = "A1"

    @Option(names = ["-i", "--in"], description = ["input file"])
    var inputFile: File? = null

    @Option(names = ["-o", "--out"], description = ["output file"])
    var outputFile: File? = null

    @Option(names = ["--head"], description = ["Reads only the first lines of input", "or (if negative) skips the initial lines"], defaultValue = "0")
    var head: Int = 0

    @Option(names = ["--tail"], description = ["Reads only the last lines of input", "or (if negative) skips the final lines"], defaultValue = "0")
    var tail: Int = 0

    @Parameters(paramLabel = "columns", description = ["Your output columns"], arity="0..*")
    var columns : List<String>? = null

    override fun run() {
        log.info("tab: $tab")
        log.info("cell: $cell")
        log.info("input file: ${inputFile?:"stdin"}")
        log.info("output file: ${outputFile?:"stdout"}")
        repl(inStream(inputFile), outStream(outputFile), tab, cell)
    }
}

/**
 *
 * cat file.csv | xcl create --template template.xsl --line-separator="\n" \
 * --field-separator="," --cell=B1 --tab=tab02 >> result.xls
 */
@Command(name = "create", mixinStandardHelpOptions = true)
class Create : Runnable {
    override fun run() {
        TODO("Not yet implemented")
    }

}

/**
 *
 * cat file.csv | xcl create --template template.xsl --line-separator="\n" \
 * --field-separator="," --cell=B1 --tab=tab02 >> result.xls
 */
@Command(name = "update", mixinStandardHelpOptions = true)
class Update : Runnable {
    override fun run() {
        TODO("Not yet implemented")
    }

}