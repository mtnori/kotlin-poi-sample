package com.mtnori.poisample.controllers

import com.mtnori.poisample.models.Greeting
import org.springframework.web.bind.annotation.GetMapping
import org.springframework.web.bind.annotation.RequestParam
import org.springframework.web.bind.annotation.RestController

@RestController
class GreetingController {

    @GetMapping("/hello")
    fun hello(@RequestParam(value = "name", required = false, defaultValue = "world") name: String): Greeting {
        return Greeting("Hello $name.")
    }
}