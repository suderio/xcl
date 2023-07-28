package net.technearts.xcl

import io.smallrye.config.ConfigMapping
import io.smallrye.config.WithName


@ConfigMapping(prefix = "greeting")
interface XCLConfig {

    @WithName("message")
    fun message(): String?

}