package com.jg.poiet.policy;

import com.jg.poiet.NiceXSSFWorkbook;
import com.jg.poiet.XSSFTemplate;
import com.jg.poiet.data.TextRenderData;
import com.jg.poiet.template.cell.CellTemplate;
import com.jg.poiet.util.StyleUtils;
import org.apache.poi.xssf.usermodel.XSSFCell;

public class TextRenderPolicy extends AbstractRenderPolicy<Object> {

    private static CellTemplate tmp ;

    @Override
    public void doRender(CellTemplate cellTemplate, Object renderData, XSSFTemplate template) {
        XSSFCell cell = cellTemplate.getCell();
        tmp = cellTemplate;
        Helper.renderTextCell(cell, renderData, template);
    }

    public static class Helper {

        public static void renderTextCell(XSSFCell cell, Object renderData,XSSFTemplate template) {
            renderTextCell(cell, renderData, template.getXSSFWorkbook());
        }

        public static void renderTextCell(XSSFCell cell, Object renderData, NiceXSSFWorkbook workbook) {
            if (null == renderData) {
                renderData = new TextRenderData();
            }
            // text
            TextRenderData textRenderData = renderData instanceof TextRenderData
                    ? (TextRenderData) renderData
                    : new TextRenderData(renderData.toString());

            String data = null == textRenderData.getText() ? "" : textRenderData.getText();

            StyleUtils.styleCell(workbook, cell, textRenderData.getStyle());

            String cellVal = cell.getRichStringCellValue().getString();
            String source = tmp.getSource();

            // if (cellVal.lastIndexOf("{{")>0) {
                cell.setCellValue(cellVal.replace(source,data));
            /*}else{
                switch (textRenderData.getDataType()) {
                    case String:
                        cell.setCellValue(data);
                        break;
                    case Double:
                        cell.setCellValue(data);
                        break;
                    default:
                        cell.setCellValue(data);
                }
            }*/
        }

    }

}
