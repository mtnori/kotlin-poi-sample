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
            excelData.writeCell("cell", 0, 0)
            excelData.writeCell("newCell", 3, 4)
            excelData.writeCellFormula("MOD(10, 3)", 2, 2)
            excelData.writeCellByName("名前定義", "cell_name")
            documentRepository.save(excelData, "output.xlsx")
        }
    }
}