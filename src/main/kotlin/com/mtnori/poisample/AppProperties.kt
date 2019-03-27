package com.mtnori.poisample

import org.springframework.boot.context.properties.ConfigurationProperties
import org.springframework.stereotype.Component

@Component
@ConfigurationProperties(prefix="app")
class AppProperties (
        var outputDir: String = ""
)