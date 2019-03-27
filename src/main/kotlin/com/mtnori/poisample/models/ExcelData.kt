package com.mtnori.poisample.models

import org.apache.poi.ss.usermodel.*
import org.apache.poi.ss.util.CellReference
import java.io.OutputStream

/**
 * Excelデータを表現するモデルクラス
 */
class ExcelData {
    private val workbook: Workbook
    private var sheet: Sheet

    /**
     * @param workbook ワークブック
     * @param sheetIdx シート番号
     */
    constructor(workbook: Workbook, sheetIdx: Int = 0) {
        this.workbook = workbook
        this.sheet = workbook.getSheetAt(sheetIdx)
    }

    /**
     * ワークブックを保存する
     * @param outputStream 出力ストリーム
     */
    fun save(outputStream: OutputStream) {
        this.workbook.write(outputStream)
    }

    /**
     * 指定した行を返す。指定した行が存在しなければ作成して返す。
     * @param rowIdx 行番号
     */
    private fun getRow(rowIdx: Int): Row {
        val row: Row? = this.sheet.getRow(rowIdx)
        if (row != null) {
            return row
        }
        return this.sheet.createRow(rowIdx)
    }

    /**
     * 指定した行を返す。指定した行が存在しなければ作成して返す。
     * @param rowIdx 行番号
     * @param colIdx 列番号
     */
    private fun getCell(rowIdx: Int, colIdx: Int): Cell {
        val row: Row = getRow(rowIdx)
        val cell: Cell? = row.getCell(colIdx);
        if (cell != null) {
            return cell
        }
        return row.createCell(colIdx)
    }

    /**
     * セルに文字列を設定する
     * @param value 文字列
     * @param rowIdx 行番号
     * @param colIdx 列番号
     */
    fun writeCell(value: String, rowIdx: Int, colIdx: Int) {
        // セルを取得する
        val cell: Cell = getCell(rowIdx, colIdx)
        // セルの値を設定する
        cell.setCellValue(value)
    }

    /**
     * 名前定義からセルを探して文字列をセットする
     */
    fun writeCellByName(value: String, name: String) {
        val name: Name? = this.workbook.getName(name)
        if (name !== null) {
            val ref: CellReference = CellReference(name.refersToFormula)
            writeCell(value, ref.row, ref.col.toInt())
        }
    }

    /**
     * セルに計算式を設定する
     * @param value 計算式
     * @param rowIdx 行番号
     * @param colIdx 列番号
     */
    fun writeCellFormula(value: String, rowIdx: Int, colIdx: Int) {
        // セルを取得する
        val cell: Cell = getCell(rowIdx, colIdx)
        // セルの値を設定する
        cell.cellFormula = value
    }
}