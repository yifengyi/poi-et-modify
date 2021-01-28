package com.jg.poiet.policy;

import com.jg.poiet.NiceXSSFWorkbook;
import com.jg.poiet.XSSFTemplate;
import com.jg.poiet.data.CellRenderData;
import com.jg.poiet.data.ListRenderData;
import com.jg.poiet.data.RowRenderData;
import com.jg.poiet.template.cell.CellTemplate;
import com.jg.poiet.util.StyleUtils;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;

/**
 * 列表处理
 */
public class ListRenderPolicy extends AbstractRenderPolicy<ListRenderData> {

    @Override
    protected boolean validate(ListRenderData data) {
        if (data.isEmpty()) {
            logger.debug("Empty ListRenderData datamodel: {}", data);
            return false;
        }
        return true;
    }

    @Override
    public void doRender(CellTemplate cellTemplate, ListRenderData data, XSSFTemplate template) {
        XSSFCell cell = cellTemplate.getCell();
        Helper.renderList(template, cell, data);
    }

    public static class Helper {

        /**
         * 渲染列表数据
         * @param template template
         * @param cell  cell
         * @param listData listData
         */
        public static void renderList(XSSFTemplate template, XSSFCell cell, ListRenderData listData) {
            if (listData.getDirection() == ListRenderData.DIRECTION.HORIZONTAL) {   //水平方向
                renderHorizontalList(template, cell, listData);
            } else if (listData.getDirection() == ListRenderData.DIRECTION.VERTICAL) {  //垂直方向
                renderVerticalList(template, cell, listData);
            }
        }

        /**
         * 渲染水平方向列表数据
         * @param template  template
         * @param cell  cell
         * @param listData  listData
         */
        private static void renderHorizontalList(XSSFTemplate template, XSSFCell cell, ListRenderData listData) {
            NiceXSSFWorkbook workbook = template.getXSSFWorkbook();
            XSSFSheet sheet = cell.getSheet();
            cell.setCellValue("");  //将该单元格值置空
            if (listData.isEmpty()) {
                return ;
            }
            insertColumn(workbook, cell, listData); //插入列
            RowRenderData rowRenderData = new RowRenderData(listData.getCellDatas());
            PolicyHelper.renderRow(template, cell.getRow(), rowRenderData, cell.getColumnIndex(), listData.getStyle());
        }

        /**
         * 插入列
         * @param workbook  workbook
         * @param cell  cell
         * @param listData  listData
         */
        private static void insertColumn(NiceXSSFWorkbook workbook, XSSFCell cell, ListRenderData listData) {
            if (!listData.isExtension()) {
                return ;
            }
            int insertNum = 0;
            for (CellRenderData cellRenderData : listData.getCellDatas()) {
                if (cellRenderData == null) {
                    continue;
                }
                insertNum += cellRenderData.getColspan() + 1;
            }
            PolicyHelper.insertColumn(workbook, cell, insertNum - 1);
        }

        /**
         * 渲染垂直方向列表
         * @param template  template
         * @param cell  cell
         * @param listData  listData
         */
        private static void renderVerticalList(XSSFTemplate template, XSSFCell cell, ListRenderData listData) {
            NiceXSSFWorkbook workbook = template.getXSSFWorkbook();
            XSSFSheet sheet = cell.getSheet();
            cell.setCellValue("");  //将该单元格值置空
            if (listData.isEmpty()) {
                return ;
            }
            insertRow(workbook, cell, listData);   //插入行
            int columnIndex = cell.getColumnIndex();
            int rowIndex = cell.getRowIndex();
            for (int i = 0; i < listData.size(); i++) {
                CellRenderData cellRenderData = listData.getCellDatas().get(i);
                XSSFRow row = sheet.getRow(rowIndex);
                if (row == null) {
                    row = sheet.createRow(rowIndex);
                }
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
                    StyleUtils.styleCell(template, xssfCell, listData.getStyle());
                    TextRenderPolicy.Helper.renderTextCell(xssfCell, cellRenderData.getRenderData(), template);
                    rowIndex += cellRenderData.getRowspan() + 1;
                } else {
                    if (cellAddress.getFirstRow() == xssfCell.getRowIndex()
                            && cellAddress.getFirstColumn() == xssfCell.getColumnIndex()) {
                        StyleUtils.styleCell(template, xssfCell, listData.getStyle());
                        TextRenderPolicy.Helper.renderTextCell(xssfCell, cellRenderData.getRenderData(), template);
                    } else {
                        i --;
                    }
                    rowIndex = cellAddress.getLastRow() + 1;
                }
            }
            template.getXSSFWorkbook().updateCellRangeAddress();
        }

        /**
         * 插入行
         * @param workbook  workbook
         * @param cell  cell
         * @param listData listData
         */
        private static void insertRow(NiceXSSFWorkbook workbook, XSSFCell cell, ListRenderData listData) {
            if (!listData.isExtension()) {
                return ;
            }
            int insertNum = 0;
            for (CellRenderData cellRenderData : listData.getCellDatas()) {
                if (cellRenderData == null) {
                    continue;
                }
                insertNum += cellRenderData.getRowspan() + 1;
            }
            PolicyHelper.insertRow(workbook, cell, insertNum - 1);
        }
    }

}
