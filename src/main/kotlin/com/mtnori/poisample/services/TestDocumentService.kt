package com.mtnori.poisample.services

import com.mtnori.poisample.repositories.DocumentRepository
import org.springframework.beans.factory.annotation.Autowired
import org.springframework.stereotype.Service

@Service
class TestDocumentService {

    @Autowired
    lateinit var documentRepository: DocumentRepository

    fun create() {
        val excelData = documentRepository.load("test.xlsx")
        if (excelData !== null) {
            excelData.print()
            documentRepository.save(excelData)
        }
    }
}