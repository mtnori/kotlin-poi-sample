package com.mtnori.poisample.controllers

import com.mtnori.poisample.controllers.dtos.RequestData
import com.mtnori.poisample.controllers.dtos.ResponseData
import com.mtnori.poisample.services.TestDocumentService
import org.springframework.web.bind.annotation.GetMapping
import org.springframework.web.bind.annotation.PostMapping
import org.springframework.web.bind.annotation.RequestBody
import org.springframework.web.bind.annotation.RestController

@RestController
class TestDocumentController(
        private val service: TestDocumentService
) {
    @GetMapping("/create")
    fun createGet(): String {
        // Excelファイルを生成する
        service.create()
        return "ok"
    }

    @PostMapping("/create")
    fun createPost(@RequestBody requestData: RequestData): ResponseData {
        println(requestData.toString())

        // Excelファイルを生成する
        val filepath = service.create2(requestData.stringData, requestData.longData)
        return ResponseData(filepath)
    }
}