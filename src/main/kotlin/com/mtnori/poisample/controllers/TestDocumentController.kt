package com.mtnori.poisample.controllers

import com.mtnori.poisample.services.TestDocumentService
import org.springframework.beans.factory.annotation.Autowired
import org.springframework.web.bind.annotation.GetMapping
import org.springframework.web.bind.annotation.RestController

@RestController
class TestDocumentController {

    @Autowired
    lateinit var service: TestDocumentService

    @GetMapping("/create")
    fun load(): String {
        // Excelファイルを生成する
        service.create()
        return "ok"
    }
}