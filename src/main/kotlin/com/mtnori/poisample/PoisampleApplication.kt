package com.mtnori.poisample

import org.springframework.boot.autoconfigure.SpringBootApplication
import org.springframework.boot.runApplication

@SpringBootApplication
class PoisampleApplication

fun main(args: Array<String>) {
    runApplication<PoisampleApplication>(*args)
}
