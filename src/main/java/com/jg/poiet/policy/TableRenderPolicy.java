package com.jg.poiet.policy;

import com.jg.poiet.NiceXSSFWorkbook;
import com.jg.poiet.XSSFTemplate;
import com.jg.poiet.data.RowRenderData;
import com.jg.poiet.data.TableRenderData;
import com.jg.poiet.template.cell.CellTemplate;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;

import java.util.Iterator;
import java.util.List;

/**
 * 表格处理
 */
public class TableRenderPolicy extends AbstractRenderPolicy<TableRenderData> {

    @Override
    protected boolean validate(TableRenderData data) {
        if (!(data).isSetBody() && !(data).isSetHeader()) {
            logger.debug("Empty MiniTableRenderData datamodel: {}", data);
            return false;
        }
        return true;
    }

    @Override
    public void doRender(CellTemplate cellTemplate, TableRenderData data, XSSFTemplate template) {
        XSSFCell cell = cellTemplate.getCell();
        Helper.renderTable(template, cell, data);
    }

    public static class Helper {
        /**
         * 渲染表格数据
         * @param template template
         * @param cell  cell
         * @param tableData tableData
         */
        public static void renderTable(XSSFTemplate template, XSSFCell cell, TableRenderData tableData) {
            NiceXSSFWorkbook workbook = template.getXSSFWorkbook();
            XSSFSheet sheet = cell.getSheet();
            cell.setCellValue("");  //将该单元格值置空
            insertRow(workbook, cell, tableData);   //插入行
            int rowIndex = cell.getRowIndex();
            if (tableData.isSetHeader()) {
                List<RowRenderData> headerData = tableData.getHeader();
                for (RowRenderData rowRenderData : headerData) {
                    if (null != rowRenderData) {
                        PolicyHelper.renderRow(template, sheet.getRow(rowIndex), rowRenderData, cell.getColumnIndex(), tableData.getHeaderStyle());
                        rowIndex ++;
                    }
                }
            }
            if (tableData.isSetBody()) {
                List<RowRenderData> bodyData = tableData.getRowDatas();
                Iterator<RowRenderData> iterator = bodyData.iterator();
                for (RowRenderData rowRenderData : bodyData) {
                    if (null != rowRenderData) {
                        PolicyHelper.renderRow(template, sheet.getRow(rowIndex), rowRenderData, cell.getColumnIndex(), tableData.getBodyStyle());
                        rowIndex ++;
                    }
                }
            }
        }

        /**
         * 插入行
         * @param workbook  workbook
         * @param cell  cell
         * @param tableData tableData
         */
        private static void insertRow(NiceXSSFWorkbook workbook, XSSFCell cell, TableRenderData tableData) {
            int insertNum = 0;
            if (tableData.isSetHeader()) {
                insertNum += tableData.getHeader().size();
            }
            if (tableData.isSetBody()) {
                insertNum += tableData.getRowDatas().size();
            }
            if (insertNum <= 1) {
                return ;
            }
            PolicyHelper.insertRow(workbook, cell, insertNum - 1);
        }
    }

}
