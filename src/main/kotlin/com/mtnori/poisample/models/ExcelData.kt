package com.mtnori.poisample.models

import org.apache.poi.ss.usermodel.*
import org.apache.poi.ss.util.CellRangeAddress
import org.apache.poi.ss.util.CellReference
import java.io.OutputStream

/**
 * Excelデータを表現するモデルクラス
 */
class ExcelData (private val workbook: Workbook, sheetIdx: Int = 0) {
    private var sheet: Sheet = workbook.getSheetAt(sheetIdx)

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
        val cell: Cell? = row.getCell(colIdx)
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
        val lName = this.workbook.getName(name)
        if (lName !== null) {
            val ref = CellReference(lName.refersToFormula)
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

    /**
     * セルを結合する
     * @param startRowIdx 開始行
     * @param mergeRowNum 行結合数
     * @param startColIdx 開始列
     * @param mergeColNum 列結合数
     */
    fun mergeCells(startRowIdx: Int, mergeRowNum: Int, startColIdx: Int, mergeColNum: Int) {
        this.sheet.addMergedRegion(CellRangeAddress(
                startRowIdx,
                startRowIdx + mergeRowNum - 1,
                startColIdx,
                startColIdx + mergeColNum - 1))
    }

    /**
     * セルの結合を解除する
     * この処理は遅いので、mergeCells関数内に組み込まないことにする
     * @param startRowIdx 開始行
     * @param mergeRowNum 行結合数
     * @param startColIdx 開始列
     * @param mergeColNum 列結合数
     */
    fun unmergeCells(startRowIdx: Int, mergeRowNum: Int, startColIdx: Int, mergeColNum: Int) {
        val targetRange = CellRangeAddress(
                startRowIdx,
                startRowIdx + mergeRowNum - 1,
                startColIdx,
                startColIdx + mergeColNum - 1)
        val mergeList: List<CellRangeAddress> = this.sheet.mergedRegions
        val removeIndices: MutableList<Int> = mutableListOf()
        for ((mergeIndex, mergedAddress) in mergeList.withIndex()) {
            if (targetRange.intersects(mergedAddress)) {
                removeIndices.add(mergeIndex)
            }
        }
        // 手前から解除するとインデックスが合わなくなるので後ろから消す
        removeIndices.reverse()
        for (removeIndex in removeIndices) {
            this.sheet.removeMergedRegion(removeIndex)
        }
    }

    /**
     * セルの計算式の計算結果を取得する
     * @param cell セル
     * @return セルの計算式の計算結果(文字列
     */
    private fun getStringFormulaValue(cell: Cell): String {
        val helper: CreationHelper = this.workbook.creationHelper
        val evaluator: FormulaEvaluator = helper.createFormulaEvaluator()
        val value: CellValue = evaluator.evaluate(cell)
        when(value.cellTypeEnum) {
            CellType.STRING -> {
                return cell.stringCellValue
            }
            CellType.NUMERIC -> {
                if (DateUtil.isCellDateFormatted(cell)) {
                    return cell.dateCellValue.toString()
                }
                return cell.numericCellValue.toString()
            }
            CellType.BOOLEAN -> {
                return cell.booleanCellValue.toString()
            }
            CellType.ERROR -> {
                val errorCode = cell.errorCellValue
                val error: FormulaError = FormulaError.forInt(errorCode)
                return error.string
            }
            CellType.BLANK -> {
                return getStringRangeValue(cell)
            }
            CellType._NONE -> {
                return ""
            }
            else -> {
                return ""
            }
        }
    }

    /**
     * 結合セルの値を取得する
     * 結合セルの場合、値が入っているのは一番左上のセルになる
     * @param cell セル
     * @return セルの値(文字列)
     */
    private fun getStringRangeValue(cell: Cell): String {
        val rowIdx = cell.rowIndex
        val colIdx = cell.columnIndex
        val size: Int = sheet.numMergedRegions
        for (i in 0..size) {
            val range: CellRangeAddress = sheet.getMergedRegion(i)
            if (range.isInRange(rowIdx, colIdx)) {
                return getCellValue(range.firstRow, range.firstColumn)
            }
        }
        return ""
    }

    /**
     * セルの値を文字列で返す
     * @param rowIdx 行番号
     * @param colIdx 列番号
     * @return セルの値(文字列)
     */
    fun getCellValue(rowIdx: Int, colIdx: Int): String {
        val cell: Cell = getCell(rowIdx, colIdx)
        when(cell.cellTypeEnum) {
            CellType.STRING -> {
                return cell.stringCellValue
            }
            CellType.NUMERIC -> {
                if (DateUtil.isCellDateFormatted(cell)) {
                    return cell.dateCellValue.toString()
                }
                return cell.numericCellValue.toString()
            }
            CellType.BOOLEAN -> {
                return cell.booleanCellValue.toString()
            }
            CellType.FORMULA -> {
                return getStringFormulaValue(cell)
            }
            CellType.ERROR -> {
                val errorCode = cell.errorCellValue
                val error: FormulaError = FormulaError.forInt(errorCode)
                return error.string
            }
            CellType.BLANK -> {
                return getStringRangeValue(cell)
            }
            CellType._NONE -> {
                return ""
            }
            else -> {
                return ""
            }
        }
    }

    /**
     * 行を挿入する
     * その際、スタイルとセル結合情報を開始行からコピーする
     * @param startRowIdx 開始行
     */
    fun insertRow(startRowIdx: Int) {
        // コピー元の行を取得
        val originalRow = getRow(startRowIdx)

        // シート末尾に新しい行を追加
        val lastRowNum = this.sheet.lastRowNum
        this.sheet.createRow(lastRowNum + 1)

        // 行を追加する位置以降の行を下にずらす
        this.sheet.shiftRows(startRowIdx + 1, lastRowNum, 1)

        // 追加した行にスタイルを設定
        val newRow = getRow(startRowIdx + 1)
        for(colIdx in 0..(originalRow.lastCellNum - 1)) {
            val originalCell = originalRow.getCell(colIdx)
            val newCell = newRow.createCell(colIdx)

            // セルのスタイルをコピー
            val newCellStyle = this.workbook.createCellStyle()
            newCellStyle.cloneStyleFrom(originalCell.cellStyle)
            newCell.cellStyle = newCellStyle

            // セルタイプのコピー
            newCell.setCellType(originalCell.cellTypeEnum)

            // 追加した行にマージ状態を設定
            if(this.sheet.numMergedRegions > 0) {
                for (index in 0..(this.sheet.numMergedRegions - 1)) {
                    val cellRangeAddress = this.sheet.getMergedRegion(index)
                    // コピー元セル(結合済み)の左上セルがmergedRegionと一致すれば、
                    // 新しいセルに対してmergeRegionを追加する
                    if (cellRangeAddress.firstRow == originalCell.rowIndex
                            && cellRangeAddress.firstColumn == originalCell.columnIndex
                    ) {
                        this.sheet.addMergedRegion((CellRangeAddress(
                                newRow.rowNum,
                                newRow.rowNum + (cellRangeAddress.lastRow - cellRangeAddress.firstRow),
                                cellRangeAddress.firstColumn,
                                cellRangeAddress.lastColumn
                        )))
                        break
                    }
                }
            }
        }
    }
}