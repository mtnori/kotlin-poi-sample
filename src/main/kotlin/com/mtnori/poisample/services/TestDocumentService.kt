package com.mtnori.poisample.services

interface TestDocumentService {
    /**
     * Excelファイルを生成する
     */
    fun create()

    /**
     * Excelファイルを作成する
     * @param stringData
     * @param longData
     * @return ファイルパス
     */
    fun create2(stringData: String, longData: Long): String
}