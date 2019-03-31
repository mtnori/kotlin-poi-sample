package com.mtnori.poisample.repositories

import com.mtnori.poisample.models.ExcelData

interface DocumentRepository {
    /**
     * 雛形ファイルを読み込む
     * @param format 雛形パス
     * @param sheetIdx シート番号
     * @return Excelデータ
     */
    fun load(format: String, sheetIdx: Int = 0): ExcelData?

    /**
     * ファイルを保存する
     * @param excelData Excelデータ
     * @param filename ファイル名
     * @return 作成したファイルのパス
     */
    fun save(excelData: ExcelData, filename: String): String
}