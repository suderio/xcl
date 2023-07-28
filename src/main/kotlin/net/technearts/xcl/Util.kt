package net.technearts.xcl

import java.io.BufferedReader
import java.io.InputStreamReader
import java.io.Reader
import java.io.StringReader

fun inPiper(input: String?): Reader {
    if (input != null)
        return StringReader(input)
    return InputStreamReader(System.`in`)
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
