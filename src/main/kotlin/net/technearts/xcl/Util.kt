package net.technearts.xcl

import java.io.*

fun inStream(input: File?): Reader {
    if (input != null)
        return FileReader(input)
    return InputStreamReader(System.`in`)
}

fun outStream(output: File?): Writer {
    if (output != null)
        return FileWriter(output)
    return OutputStreamWriter(System.out);
}

fun repl(input: Reader, output: Writer, tab: String, cell: String) {
    BufferedWriter(output).use { writer ->
        BufferedReader(input).use { reader ->
            reader.lines().forEach(writer::appendLine)
        }
    }
}


fun Reader.string(): String {
    while (true) {
        BufferedReader(this).use { reader ->
            var inputStr: String? = null
            if (reader.readLine().also { inputStr = it } != null) {
                return inputStr ?: ""
            }
        }
    }
}
