package com.mtnori.poisample.repositories.impl

import com.mtnori.poisample.AppProperties
import com.mtnori.poisample.models.ExcelData
import com.mtnori.poisample.repositories.DocumentRepository
import org.apache.poi.openxml4j.exceptions.InvalidFormatException
import org.apache.poi.ss.usermodel.Workbook
import org.apache.poi.ss.usermodel.WorkbookFactory
import org.springframework.core.io.ClassPathResource
import org.springframework.stereotype.Repository
import java.io.FileOutputStream
import java.io.IOException
import java.io.InputStream
import java.io.OutputStream

@Repository
class DocumentRepositoryImpl(
        private val appProperties: AppProperties
): DocumentRepository  {
    /**
     * 雛形ファイルを読み込む
     * @param format 雛形パス
     * @param sheetIdx シート番号
     * @return Excelデータ
     */
    override fun load(format: String, sheetIdx: Int): ExcelData? {
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

    /**
     * ファイルを保存する
     * @param excelData Excelデータ
     * @param filename ファイル名
     */
    override fun save(excelData: ExcelData, filename: String) {
        var outputStream: OutputStream? = null
        try {
            outputStream = FileOutputStream("${appProperties.outputDir}/$filename")
            excelData.save(outputStream)
        } catch (e: IOException) {
            throw e
        } finally {
            try {
                outputStream?.close()
            } catch (e: IOException) {
                throw e
            }
        }
    }
}