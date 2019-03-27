package com.mtnori.poisample.models

import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.ss.usermodel.Workbook

/**
 * Excelデータを表現するモデルクラス
 */
class ExcelData {
    val workbook: Workbook
    private var sheet: Sheet

    /**
     * @param workbook ワークブック
     * @param sheetIdx シート番号
     */
    constructor(workbook: Workbook, sheetIdx: Int = 0) {
        this.workbook = workbook
        this.sheet = workbook.getSheetAt(sheetIdx)
    }

    fun print() {
        this.sheet.forEach {
            it.forEach {
                if (it.columnIndex > 0) print(",")
                print(it)
            }
            println()
        }
    }
}