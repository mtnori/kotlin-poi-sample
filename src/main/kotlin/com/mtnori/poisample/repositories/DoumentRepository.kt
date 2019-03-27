package com.mtnori.poisample.repositories

import com.fasterxml.jackson.databind.exc.InvalidFormatException
import com.mtnori.poisample.AppProperties
import com.mtnori.poisample.models.ExcelData
import org.apache.poi.ss.usermodel.Workbook
import org.apache.poi.ss.usermodel.WorkbookFactory
import org.springframework.beans.factory.annotation.Autowired
import org.springframework.core.io.ClassPathResource
import org.springframework.stereotype.Repository
import java.io.FileOutputStream
import java.io.IOException
import java.io.InputStream

@Repository
class DocumentRepository {

    @Autowired
    lateinit var appProperties: AppProperties

    /**
     * 雛形ファイルを読み込む
     * @param format 雛形パス
     * @param sheetIdx シート番号
     * @return 雛形ファイル情報
     */
    fun load(format: String, sheetIdx: Int = 0): ExcelData? {
        var inputStream: InputStream? = null
        val workbook: Workbook?
        try {
            val resource = ClassPathResource("formats/$format")
            inputStream = resource.inputStream
            workbook = WorkbookFactory.create(inputStream)
        } catch (e: IOException) {
            throw e
        } catch (e: InvalidFormatException) {
            throw e
        } finally {
            try {
                inputStream?.close()
            } catch (e: IOException){
                throw e
            }
        }
        if (workbook != null) {
            return ExcelData(workbook, sheetIdx)
        }
        return null
    }

    fun save(excelData: ExcelData) {
        val outputStream = FileOutputStream("${appProperties.outputDir}/output.xlsx")
        excelData.workbook.write(outputStream)
        outputStream.close()
    }
}