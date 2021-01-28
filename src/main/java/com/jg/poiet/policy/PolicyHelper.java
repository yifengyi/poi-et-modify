package com.jg.poiet.policy;

import com.jg.poiet.NiceXSSFWorkbook;
import com.jg.poiet.XSSFTemplate;
import com.jg.poiet.data.CellRenderData;
import com.jg.poiet.data.RowRenderData;
import com.jg.poiet.data.style.Style;
import com.jg.poiet.util.StyleUtils;
import org.apache.poi.ss.usermodel.CellCopyPolicy;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;

import java.util.List;

/**
 * policy助手类
 */
public class PolicyHelper {

    private PolicyHelper() {
        throw new RuntimeException("当前类无法被实例化！");
    }

    /**
     * 插入行(包含当前行数据)
     * @param workbook  workbook
     * @param cell  cell
     * @param insertNum insertNum
     */
    public static void insertRow(NiceXSSFWorkbook workbook, XSSFCell cell, int insertNum) {
        XSSFSheet sheet = cell.getSheet();
        workbook.insertRowsAfter(workbook.getSheetIndex(sheet), cell.getRowIndex(), insertNum);
        for (int i = cell.getRowIndex() + 1; i <= cell.getRowIndex() + insertNum; i++) {
            XSSFRow row = sheet.createRow(i);
            row.copyRowFrom(cell.getRow(), new CellCopyPolicy());
        }
        XSSFRow xssfRow = cell.getRow();
        for(int i = 0; i <= xssfRow.getLastCellNum(); i++){ //将当前行所在的合并单元格和扩展出去的行合并
            XSSFCell xssfCell = xssfRow.getCell(i);
            if (null == xssfCell) continue;
            CellRangeAddress cellAddress = workbook.getCellRangeAddress(xssfCell);
            if (null == cellAddress) continue;
            if (xssfCell.getColumnIndex() != cellAddress.getFirstColumn()) continue;
            if (cellAddress.getFirstRow() == cellAddress.getLastRow()) continue;
            if (cellAddress.getLastRow() != cell.getRowIndex()) continue;
            int firstRow = cellAddress.getFirstRow();
            int lastRow = cellAddress.getLastRow() + insertNum;
            int firstColumn = cellAddress.getFirstColumn();
            int lastColumn = cellAddress.getLastColumn();
            workbook.removeMergedRegion(xssfCell, false);
            workbook.addMergedRegion(workbook.getSheetIndex(sheet), firstRow, lastRow, firstColumn, lastColumn);
        }
    }

    /**
     * 插入列(包含当前列数据)
     * @param workbook  workbook
     * @param cell  cell
     * @param insertNum insertNum
     */
    public static void insertColumn(NiceXSSFWorkbook workbook, XSSFCell cell, int insertNum) {
        XSSFSheet sheet = cell.getSheet();
        workbook.insertColumnsAfter(workbook.getSheetIndex(sheet), cell.getColumnIndex(), insertNum);
        for (int i = 0; i <= sheet.getLastRowNum(); i++) {  //写入数据
            XSSFRow row = sheet.getRow(i);
            if (row == null) {
                continue;
            }
            XSSFCell sourceCell = row.getCell(cell.getColumnIndex());
            if (sourceCell == null) {
                continue;
            }
            if (workbook.getCellRangeAddress(sourceCell) != null) {
                continue;
            }
            for (int j = cell.getColumnIndex() + insertNum; j > cell.getColumnIndex(); j--) {
                XSSFCell targetCell = row.getCell(j);
                targetCell.copyCellFrom(sourceCell, new CellCopyPolicy());
            }
        }
    }

    /**
     * 渲染行数据
     * @param template  template
     * @param row   row
     * @param rowRenderData rowRenderData
     * @param columnIndex   columnIndex
     * @param style style
     */
    public static void renderRow(XSSFTemplate template, XSSFRow row, RowRenderData rowRenderData, int columnIndex, Style style) {
        NiceXSSFWorkbook workbook = template.getXSSFWorkbook();
        List<CellRenderData> cellRenderDataList = rowRenderData.getCellDatas();
        for (int i = 0; i < cellRenderDataList.size(); i++) {
            CellRenderData cellRenderData = cellRenderDataList.get(i);
            if (null == cellRenderData || null == cellRenderData.getRenderData()) continue;
            XSSFCell xssfCell = row.getCell(columnIndex);
            if (null == xssfCell) {
                xssfCell = row.createCell(columnIndex);
            }
            CellRangeAddress cellAddress = workbook.getCellRangeAddress(xssfCell);
            if (null == cellAddress) {
                if (cellRenderData.getRowspan() > 0 || cellRenderData.getColspan() > 0) {   //需要合并单元格
                    workbook.addMergedRegion(template.getXSSFWorkbook().getSheetIndex(row.getSheet()),
                            xssfCell.getRowIndex(), xssfCell.getRowIndex() + cellRenderData.getRowspan(),
                            xssfCell.getColumnIndex(), xssfCell.getColumnIndex() + cellRenderData.getColspan(), false);
                }
                StyleUtils.styleCell(template, xssfCell, style, rowRenderData.getStyle());
                TextRenderPolicy.Helper.renderTextCell(xssfCell, cellRenderData.getRenderData(), template);
                columnIndex += cellRenderData.getColspan() + 1;
            } else {
                if (cellAddress.getFirstRow() == xssfCell.getRowIndex()
                        && cellAddress.getFirstColumn() == xssfCell.getColumnIndex()) {
                    StyleUtils.styleCell(template, xssfCell, style, rowRenderData.getStyle());
                    TextRenderPolicy.Helper.renderTextCell(xssfCell, cellRenderData.getRenderData(), template);
                } else {
                    i--;
                }
                columnIndex = cellAddress.getLastColumn() + 1;
            }
        }
        template.getXSSFWorkbook().updateCellRangeAddress();
    }
}
