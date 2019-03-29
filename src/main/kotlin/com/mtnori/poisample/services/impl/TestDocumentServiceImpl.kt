package com.mtnori.poisample.services.impl

import com.mtnori.poisample.repositories.DocumentRepository
import com.mtnori.poisample.services.TestDocumentService
import org.springframework.stereotype.Service

@Service
class TestDocumentServiceImpl(
        private val documentRepository: DocumentRepository
): TestDocumentService {
      override fun create() {
        val excelData = documentRepository.load("test.xlsx")
        if (excelData !== null) {
            // 既存セルに値をセットする
            excelData.writeCell("cell", 0, 0)
            // 新規セルに値をセットする
            excelData.writeCell("newCell", 3, 4)
            // 計算式をセットする
            excelData.writeCellFormula("MOD(10, 3)", 2, 2)
            // 名前定義からセルを特定し値をセットする
            excelData.writeCellByName("名前定義", "cell_name")
            // セルの結合
            excelData.mergeCells(15, 2, 0, 3)
            // セルの結合(結合されているセルは再結合できないので、結合を解除する)
            excelData.unMergeCells(7, 2, 0, 3)
            excelData.mergeCells(7, 2, 0, 3)
            // セルの値を取得する
            println(excelData.getCellValue(19,0)) // 文字列
            println(excelData.getCellValue(20,0)) // 数値
            println(excelData.getCellValue(21,0)) // 真偽値
            println(excelData.getCellValue(22,0)) // 計算結果
            println(excelData.getCellValue(23,0)) // エラー
            println(excelData.getCellValue(25,1)) // 結合セル
            println(excelData.getCellValue(26,0)) // 日付
            documentRepository.save(excelData, "output.xlsx")
        }
        val excelData2 = documentRepository.load("test2.xlsx")
        if (excelData2 !== null) {
            // 行を挿入する
            excelData2.insertRow(3)
            documentRepository.save(excelData2, "output2.xlsx")
        }
    }
}