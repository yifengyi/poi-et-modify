package com.jg.poiet;

import com.jg.poiet.data.PictureRenderData;
import com.jg.poiet.data.TextRenderData;
import com.jg.poiet.data.style.Style;
import com.jg.poiet.exception.RenderException;
import com.jg.poiet.exception.ResolverException;
import com.jg.poiet.policy.PictureRenderPolicy;
import com.jg.poiet.policy.TextRenderPolicy;
import com.jg.poiet.util.CellUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.CellCopyPolicy;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.*;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.*;

/**
 * 对原生poi的扩展
 */
public class NiceXSSFWorkbook extends XSSFWorkbook {

    private static Logger logger = LoggerFactory.getLogger(NiceXSSFWorkbook.class);

    protected Map<String, CellRangeAddress> cellRangeAddressMap = new HashMap<>();
    protected Map<CellRangeAddress, Integer> cellRangeAddressIntegerMap = new HashMap<>();
    protected List<CellRangeAddress> allCellRangeAddress = new ArrayList<>();

    private int DEFAULT_COLUMN_WIDTH = 2340;

    /**
     * 新建空白工作簿
     */
    public NiceXSSFWorkbook() {
        super();
        this.createSheet();
    }

    /**
     * 通过流的形式打开一个工作簿
     * @param in    excel流
     * @throws IOException
     */
    public NiceXSSFWorkbook(InputStream in) throws IOException {
        super(in);
        buildAllCellRangeAddress();
    }

    /**
     * 通过文件的形式打开一个工作簿
     * @param file  excel文件
     * @return
     */
    public static NiceXSSFWorkbook compile(File file) {
        try {
            return new NiceXSSFWorkbook(new FileInputStream(file));
        } catch (IOException e) {
            logger.error("Cannot find the file", e);
            throw new ResolverException("Cannot find the file [" + file.getPath() + "]");
        }
    }

    /**
     * 通过文件路径的形式打开一个工作簿
     * @param path  excel文件路径
     * @return
     */
    public static NiceXSSFWorkbook compile(String path) {
        return NiceXSSFWorkbook.compile(new File(path));
    }

    /**
     * 将结果输出到任意流中
     * @param outputStream
     * @throws IOException
     */
    public void writeToOutputStream(OutputStream outputStream) throws IOException {
        this.write(outputStream);
        this.close();
    }

    /**
     * 将结果输出到本地文件中
     * @param path  文件路径
     * @throws IOException
     */
    public void writeToFile(String path) throws IOException {
        FileOutputStream out = new FileOutputStream(path);
        this.write(out);
        this.close();
        out.flush();
        out.close();
    }

    /**
     * 获取单元格数据
     * @param sheetIndex    工作表位置
     * @param rowIndex  行位置
     * @param columnIndex   列位置
     * @return  String类型数据
     */
    public String getStringCellValue(int sheetIndex, int rowIndex, int columnIndex) {
        XSSFCell cell = this.getCell(sheetIndex, rowIndex, columnIndex);
        return CellUtils.getCellValue(cell);
    }

    /**
     * 获取单元格数据
     * @param sheetIndex    工作表位置
     * @param rowIndex  行位置
     * @param columnIndex   列位置
     * @return  Double类型数据
     */
    public Double getDoubleCellValue(int sheetIndex, int rowIndex, int columnIndex) {
        String cellValue = this.getStringCellValue(sheetIndex, rowIndex, columnIndex);
        if (StringUtils.isEmpty(cellValue)) {
            return null;
        }
        return Double.valueOf(cellValue);
    }

    /**
     * 获取单元格数据
     * @param sheetIndex    工作表位置
     * @param rowIndex  行位置
     * @param columnIndex   列位置
     * @return  Float类型数据
     */
    public Float getFloatCellValue(int sheetIndex, int rowIndex, int columnIndex) {
        String cellValue = this.getStringCellValue(sheetIndex, rowIndex, columnIndex);
        if (StringUtils.isEmpty(cellValue)) {
            return null;
        }
        return Float.valueOf(cellValue);
    }

    /**
     * 获取单元格数据
     * @param sheetIndex    工作表位置
     * @param rowIndex  行位置
     * @param columnIndex   列位置
     * @return  Long类型数据
     */
    public Long getLongCellValue(int sheetIndex, int rowIndex, int columnIndex) {
        String cellValue = this.getStringCellValue(sheetIndex, rowIndex, columnIndex);
        if (StringUtils.isEmpty(cellValue)) {
            return null;
        }
        if (cellValue.indexOf('.') > 0) {
            cellValue = cellValue.substring(0, cellValue.indexOf('.'));
        }
        return Long.valueOf(cellValue);
    }

    /**
     * 获取单元格数据
     * @param sheetIndex    工作表位置
     * @param rowIndex  行位置
     * @param columnIndex   列位置
     * @return  Integer类型数据
     */
    public Integer getIntegerCellValue(int sheetIndex, int rowIndex, int columnIndex) {
        String cellValue = this.getStringCellValue(sheetIndex, rowIndex, columnIndex);
        if (StringUtils.isEmpty(cellValue)) {
            return null;
        }
        if (cellValue.indexOf('.') > 0) {
            cellValue = cellValue.substring(0, cellValue.indexOf('.'));
        }
        return Integer.valueOf(cellValue);
    }

    /**
     * 获取单元格数据
     * @param sheetIndex    工作表位置
     * @param rowIndex  行位置
     * @param columnIndex   列位置
     * @return  Date类型数据
     */
    public Date getDateCellValue(int sheetIndex, int rowIndex, int columnIndex) {
        try {
            String cellValue = this.getStringCellValue(sheetIndex, rowIndex, columnIndex);
            if (StringUtils.isEmpty(cellValue)) {
                return null;
            }
            DateFormat formater = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
            return formater.parse(cellValue);
        } catch (Exception e) {
            throw new RenderException(e);
        }
    }

    /**
     * 设置单元格的值
     * @param sheetIndex    工作表位置
     * @param rowIndex  行位置
     * @param columnIndex   列位置
     * @param value 设置值
     */
    public void setCellValue(int sheetIndex, int rowIndex, int columnIndex, String value) {
        setCellValue(sheetIndex, rowIndex, columnIndex, value, null);
    }

    /**
     * 设置单元格的值
     * @param sheetIndex    工作表位置
     * @param rowIndex  行位置
     * @param columnIndex   列位置
     * @param value 设置值
     * @param style 单元格样式
     */
    public void setCellValue(int sheetIndex, int rowIndex, int columnIndex, String value, Style style) {
        setCellValue(sheetIndex, rowIndex, columnIndex, new TextRenderData(value, style));
    }

    /**
     * 设置单元格的值
     * @param sheetIndex    工作表位置
     * @param rowIndex  行位置
     * @param columnIndex   列位置
     * @param textRenderData    文本渲染数据对象
     */
    public void setCellValue(int sheetIndex, int rowIndex, int columnIndex, TextRenderData textRenderData) {
        XSSFCell cell = this.getCell(sheetIndex, rowIndex, columnIndex);
        setCellValue(cell, textRenderData);
    }

    /**
     * 设置单元格的值
     * @param cell  单元格对象
     * @param value 待写入的值
     */
    public void setCellValue(XSSFCell cell, String value) {
        setCellValue(cell, value, null);
    }

    /**
     * 设置单元格的值
     * @param cell  单元格对象
     * @param value 待写入的值
     * @param style 单元格的样式对象
     */
    public void setCellValue(XSSFCell cell, String value, Style style) {
        setCellValue(cell, new TextRenderData(value, style));
    }

    /**
     * 设置单元格的值
     * @param cell  单元格对象
     * @param textRenderData 文本渲染数据对象
     */
    public void setCellValue(XSSFCell cell, TextRenderData textRenderData) {
        TextRenderPolicy.Helper.renderTextCell(cell, textRenderData, this);
    }

    /**
     * 插入图片到单元格中
     * @param sheetIndex    工作表位置
     * @param rowIndex  行位置
     * @param columnIndex   列位置
     * @param pictureRenderData 图片渲染数据对象
     */
    public void setCellValue(int sheetIndex, int rowIndex, int columnIndex, PictureRenderData pictureRenderData) {
        XSSFCell cell = this.getCell(sheetIndex, rowIndex, columnIndex);
        setCellValue(cell, pictureRenderData);
    }

    /**
     * 插入图片到单元格中
     * @param cell  单元格对象
     * @param pictureRenderData 图片渲染数据对象
     */
    public void setCellValue(XSSFCell cell, PictureRenderData pictureRenderData) {
        PictureRenderPolicy.Helper.renderPicture(cell, pictureRenderData, this);
    }

    /**
     * 构建所有的合并单元格对象
     */
    private void buildAllCellRangeAddress() {
        for (int i = 0; i < this.getNumberOfSheets(); i++) {
            XSSFSheet xssfSheet = this.getSheetAt(i);
            List<CellRangeAddress> cellRangeAddresses = xssfSheet.getMergedRegions();
            if (null == cellRangeAddresses) continue;
            allCellRangeAddress.addAll(cellRangeAddresses);
            for (int j = 0; j < xssfSheet.getNumMergedRegions(); j++) {
                CellRangeAddress cellAddress = xssfSheet.getMergedRegion(j);
                cellRangeAddressIntegerMap.put(cellAddress, j);
                int firstRow = cellAddress.getFirstRow();
                int firstColumn = cellAddress.getFirstColumn();
                for (int ri = firstRow; ri <= cellAddress.getLastRow(); ri++) {
                    for (int cj = firstColumn; cj <= cellAddress.getLastColumn(); cj++) {
                        String key = this.getCellRangeAddressMapKey(i, ri, cj);
                        cellRangeAddressMap.put(key, cellAddress);
                    }
                }
            }
        }
    }

    /**
     * 获取合并单元格map的key
     * @param sheetIndex
     * @param rowIndex
     * @param columnIndex
     * @return
     */
    private String getCellRangeAddressMapKey(int sheetIndex, int rowIndex, int columnIndex) {
        return sheetIndex + "_" + rowIndex + "_" + columnIndex;
    }

    /**
     * 更新合并单元格
     */
    public void updateCellRangeAddress() {
        cellRangeAddressMap.clear();
        allCellRangeAddress.clear();
        cellRangeAddressIntegerMap.clear();
        buildAllCellRangeAddress();
    }

    /**
     * 判断单元格是否为合并单元格
     * @param cell  单元格对象
     * @return
     */
    public boolean isMergedRegion(XSSFCell cell) {
        int sheetIndex = this.getSheetIndex(cell.getSheet());
        int rowIndex = cell.getRowIndex();
        int columnIndex = cell.getColumnIndex();
        return isMergedRegion(sheetIndex, rowIndex, columnIndex);
    }

    /**
     * 判断单元格是否为合并单元格起点
     * @param cell  单元格对象
     * @return
     */
    public boolean isMergedRegionBegin(XSSFCell cell) {
        CellRangeAddress cellAddresses = this.getCellRangeAddress(cell);
        return cellAddresses != null && cellAddresses.getFirstRow() == cell.getRowIndex() && cellAddresses.getFirstColumn() == cell.getColumnIndex();
    }

    /**
     * 判断单元格是否为合并单元格起点
     * @param sheetIndex
     * @param rowIndex
     * @param columnIndex
     * @return
     */
    public boolean isMergedRegionBegin(int sheetIndex, int rowIndex, int columnIndex) {
        CellRangeAddress cellAddresses = this.getCellRangeAddress(sheetIndex, rowIndex, columnIndex);
        return cellAddresses != null && cellAddresses.getFirstRow() == rowIndex && cellAddresses.getFirstColumn() == columnIndex;
    }

    /**
     * 判断单元格是否为合并单元格
     * @param sheetIndex
     * @param rowIndex
     * @param columnIndex
     * @return  true:表示是，false：表示否
     */
    public boolean isMergedRegion(int sheetIndex, int rowIndex, int columnIndex) {
        return getCellRangeAddress(sheetIndex, rowIndex, columnIndex) != null;
    }

    /**
     * 获取合并单元格
     * @param sheetIndex    sheet index
     * @param rowIndex  row index
     * @param columnIndex   column index
     * @return  CellRangeAddress
     */
    public CellRangeAddress getCellRangeAddress(int sheetIndex, int rowIndex, int columnIndex) {
        return cellRangeAddressMap.get(getCellRangeAddressMapKey(sheetIndex, rowIndex, columnIndex));
    }

    /**
     * 获取合并单元格
     * @param cell  单元格
     * @return  CellRangeAddress
     */
    public CellRangeAddress getCellRangeAddress(XSSFCell cell) {
        int sheetIndex = this.getSheetIndex(cell.getSheet());
        int rowIndex = cell.getRowIndex();
        int columnIndex = cell.getColumnIndex();
        return getCellRangeAddress(sheetIndex, rowIndex, columnIndex);
    }

    /**
     * 合并单元格
     * @param sheetIndex    sheetIndex
     * @param firstRowIndex firstRowIndex
     * @param lastRowIndex  lastRowIndex
     * @param firstColumnIndex  firstColumnIndex
     * @param lastColumnIndex   lastColumnIndex
     * @param isUpdate  是否更新
     */
    public void addMergedRegion(int sheetIndex, int firstRowIndex, int lastRowIndex, int firstColumnIndex, int lastColumnIndex, boolean isUpdate) {
        if (firstRowIndex > lastRowIndex || firstColumnIndex > lastColumnIndex) return ;
        if (firstRowIndex == lastRowIndex && firstColumnIndex == lastColumnIndex) return ;
        XSSFSheet sheet = this.getSheetAt(sheetIndex);
        CellRangeAddress region = new CellRangeAddress(firstRowIndex, lastRowIndex, firstColumnIndex, lastColumnIndex);
        //合并之后，将非第一个单元格的值置空,将其他单元格的样式设置成第一个单元格的样式
        for (int i = firstRowIndex; i <= lastRowIndex; i++) {
            XSSFRow xssfRow = sheet.getRow(i);
            if (xssfRow == null) {
                continue;
            }
            for (int j = firstColumnIndex; j <= lastColumnIndex; j++) {
                if (i == firstRowIndex && j == firstColumnIndex) {
                    continue;
                }
                XSSFCell cell = xssfRow.getCell(j);
                if (cell != null) {
                    cell.setCellValue("");
                }
            }
        }
        sheet.addMergedRegion(region);
        if (isUpdate) {
            updateCellRangeAddress();
        } else {
            allCellRangeAddress.add(region);
            for (int ri = firstRowIndex; ri <= lastRowIndex; ri++) {
                for (int cj = firstColumnIndex; cj <= lastColumnIndex; cj++) {
                    String key = this.getCellRangeAddressMapKey(sheetIndex, ri, cj);
                    cellRangeAddressMap.put(key, region);
                }
            }
        }
    }

    /**
     * 合并单元格
     * @param sheetIndex    sheetIndex
     * @param firstRowIndex firstRowIndex
     * @param lastRowIndex  lastRowIndex
     * @param firstColumnIndex  firstColumnIndex
     * @param lastColumnIndex   lastColumnIndex
     */
    public void addMergedRegion(int sheetIndex, int firstRowIndex, int lastRowIndex, int firstColumnIndex, int lastColumnIndex) {
        this.addMergedRegion(sheetIndex, firstRowIndex, lastRowIndex, firstColumnIndex, lastColumnIndex, true);
    }

    /**
     * 拆分单元格
     * @param cell
     */
    public void removeMergedRegion(XSSFCell cell) {
        removeMergedRegion(cell, true);
    }

    /**
     * 拆分单元格
     * @param cell   单元格
     * @param isUpdate 是否更新合并单元格信息
     */
    public void removeMergedRegion(XSSFCell cell, boolean isUpdate) {
        CellRangeAddress cellAddress = this.getCellRangeAddress(cell);
        if (null == cellAddress) return;
        cell.getSheet().removeMergedRegion(cellRangeAddressIntegerMap.get(cellAddress));
        if (isUpdate) {
            updateCellRangeAddress();
        }
    }

    /**
     * 在第sheetIndex个sheet中的第insertRowIndex行之前插入insertNum行
     * @param sheetIndex    sheetIndex
     * @param insertRowIndex    insertRowIndex
     * @param insertNum insertNum
     */
    public void insertRowsBefore(int sheetIndex, int insertRowIndex, int insertNum) {
        if (insertNum <= 0) return;
        //TO DO 插入行
        XSSFSheet sheet = this.getSheetAt(sheetIndex);
        if (null == sheet) throw new RenderException("工作表不存在！") ;
        //插入行
        sheet.createRow(sheet.getLastRowNum() + insertNum);
        for (int i = sheet.getLastRowNum(); i >= insertRowIndex + insertNum; i--) {
            XSSFRow targetRow = sheet.createRow(i);
            XSSFRow sourceRow = sheet.getRow(i - insertNum);
            this.copyRow(targetRow, sourceRow);
        }
        //清空行数据和样式（单元格除外）
        for (int i = insertRowIndex + insertNum - 1; i >= insertRowIndex; i--) {
            sheet.createRow(i);
        }
        //遍历合并单元格
        XSSFRow currentRow = sheet.getRow(insertRowIndex + insertNum);
        List<CellRangeAddress> cellRangeAddressList = new ArrayList<>(sheet.getMergedRegions());
        for (CellRangeAddress cellAddress : cellRangeAddressList) {
            int firstRow = cellAddress.getFirstRow();
            int lastRow = cellAddress.getLastRow();
            int firstColumn = cellAddress.getFirstColumn();
            int lastColumn = cellAddress.getLastColumn();
            if (firstRow < insertRowIndex && lastRow >= insertRowIndex) {

                for (int i = insertRowIndex; i <= insertRowIndex + insertNum - 1; i ++) {   //复制单元格样式
                    XSSFRow row = sheet.getRow(i);
                    if (row == null) {
                        row = sheet.createRow(i);
                    }
                    for (int j = firstColumn; j <= lastColumn; j++) {
                        XSSFCell sourceCell = currentRow.getCell(j);
                        if (sourceCell != null) {
                            XSSFCell targetCell = row.createCell(j);
                            targetCell.copyCellFrom(sourceCell, new CellCopyPolicy());
                        }
                    }
                }

                sheet.removeMergedRegion(cellRangeAddressIntegerMap.get(cellAddress));
                addMergedRegion(this.getSheetIndex(sheet),
                        firstRow,
                        lastRow + insertNum,
                        firstColumn, lastColumn);
            }
        }
        this.updateCellRangeAddress();
    }

    /**
     * 在第sheetIndex个sheet中的第insertRowIndex行之后插入insertNum行
     * @param sheetIndex    sheetIndex
     * @param insertRowIndex    insertRowIndex
     * @param insertNum insertNum
     */
    public void insertRowsAfter(int sheetIndex, int insertRowIndex, int insertNum) {
        if (insertNum <= 0) return;
        //TO DO 插入行
        XSSFSheet sheet = this.getSheetAt(sheetIndex);
        if (null == sheet) throw new RenderException("工作表不存在！") ;
        sheet.createRow(sheet.getLastRowNum() + insertNum);
        for (int i = sheet.getLastRowNum(); i > insertRowIndex + insertNum; i--) {
            XSSFRow targetRow = sheet.createRow(i);
            XSSFRow sourceRow = sheet.getRow(i - insertNum);
            this.copyRow(targetRow, sourceRow);
        }
        for (int i = insertRowIndex + insertNum; i > insertRowIndex; i--) {
            sheet.createRow(i);
        }
        //遍历合并单元格
        XSSFRow currentRow = sheet.getRow(insertRowIndex);
        List<CellRangeAddress> cellRangeAddressList = new ArrayList<>(sheet.getMergedRegions());
        for (CellRangeAddress cellAddress : cellRangeAddressList) {
            int firstRow = cellAddress.getFirstRow();
            int lastRow = cellAddress.getLastRow();
            int firstColumn = cellAddress.getFirstColumn();
            int lastColumn = cellAddress.getLastColumn();
            if (lastRow > insertRowIndex && firstRow <= insertRowIndex) {

                for (int i = insertRowIndex + 1; i <= insertRowIndex + insertNum; i ++) {   //复制单元格样式
                    XSSFRow row = sheet.getRow(i);
                    if (row == null) {
                        row = sheet.createRow(i);
                    }
                    for (int j = firstColumn; j <= lastColumn; j++) {
                        XSSFCell sourceCell = currentRow.getCell(j);
                        if (sourceCell != null) {
                            XSSFCell targetCell = row.createCell(j);
                            targetCell.copyCellFrom(sourceCell, new CellCopyPolicy());
                        }
                    }
                }

                sheet.removeMergedRegion(cellRangeAddressIntegerMap.get(cellAddress));
                addMergedRegion(this.getSheetIndex(sheet),
                        firstRow,
                        lastRow + insertNum,
                        firstColumn, lastColumn);
            }
        }
        this.updateCellRangeAddress();
    }

    /**
     * 复制行
     * @param targetRow 目标行
     * @param sourceRow 源行
     */
    public void copyRow(XSSFRow targetRow, XSSFRow sourceRow) {
        if (sourceRow == null) {
            return ;
        }
        targetRow.copyRowFrom(sourceRow, new CellCopyPolicy());
        XSSFSheet sheet = targetRow.getSheet();
        for (int i = 0; i <= sourceRow.getLastCellNum(); i++) {
            XSSFCell xssfCell = sourceRow.getCell(i);
            if (null == xssfCell) continue;
            CellRangeAddress cellAddresses = getCellRangeAddress(xssfCell);
            if (null == cellAddresses) continue;
            if (xssfCell.getRowIndex() != cellAddresses.getFirstRow()
                    || xssfCell.getColumnIndex() != cellAddresses.getFirstColumn()) continue;
            int firstRow = cellAddresses.getFirstRow();
            int firstColumn = cellAddresses.getFirstColumn();
            int lastRow = cellAddresses.getLastRow();
            int lastColumn = cellAddresses.getLastColumn();
            int regionIndex = cellRangeAddressIntegerMap.get(cellAddresses);
            sheet.removeMergedRegion(cellRangeAddressIntegerMap.get(cellAddresses));
            addMergedRegion(this.getSheetIndex(sheet),
                    targetRow.getRowNum(),
                    targetRow.getRowNum() + (lastRow - firstRow),
                    firstColumn, lastColumn);
        }
    }

    /**
     * 复制行
     * @param sheetIndex    sheetIndex
     * @param targetRowIndex    目标行位置
     * @param sourceRowIndex    源行位置
     */
    public void copyRow(int sheetIndex, int targetRowIndex, int sourceRowIndex) {
        XSSFSheet sheet = this.getSheetAt(sheetIndex);
        if (sheet == null) {
            throw new RenderException("工作表不存在！");
        }
        XSSFRow sourceRow = sheet.getRow(sourceRowIndex);
        XSSFRow targetRow = sheet.createRow(targetRowIndex);
        copyRow(targetRow, sourceRow);
    }

    /**
     * 删除行
     * @param sheetIndex    sheetIndex
     * @param rowIndex  rowIndex
     */
    public void removeRow(int sheetIndex, int rowIndex) {
        if (rowIndex < 0) return;
        XSSFSheet sheet = this.getSheetAt(sheetIndex);
        if (null == sheet) return ;

        //处理合并单元格
        List<CellRangeAddress> cellRangeAddressList = new ArrayList<>(sheet.getMergedRegions());
        for (CellRangeAddress cellAddress : cellRangeAddressList) {
            int firstRow = cellAddress.getFirstRow();
            int lastRow = cellAddress.getLastRow();
            int firstColumn = cellAddress.getFirstColumn();
            int lastColumn = cellAddress.getLastColumn();
            //以删除行为起点的合并单元格全部拆分，不保留，如果存在跨行的合并单元格，单元格全部清除。
            if (firstRow == rowIndex) {
                sheet.removeMergedRegion(cellRangeAddressIntegerMap.get(cellAddress));
                updateCellRangeAddress();
                for (int i = firstRow; i <= lastRow; i++) {
                    XSSFRow row = sheet.getRow(i);
                    if (row == null) continue;
                    for (int j = firstColumn; j <= lastColumn; j++) {
                        row.createCell(j);
                    }
                }
            }
            else if (lastRow >= rowIndex) {  //不以删除行为起点的合并单元格，合并行统一-1，样式保留。
                sheet.removeMergedRegion(cellRangeAddressIntegerMap.get(cellAddress));
                updateCellRangeAddress();
                addMergedRegion(this.getSheetIndex(sheet),
                        firstRow,
                        lastRow - 1,
                        firstColumn, lastColumn);
            }
        }

        //从删除行开始遍历，复制行。
        for (int i = rowIndex; i < sheet.getLastRowNum(); i++) {
            XSSFRow targetRow = sheet.createRow(i);
            XSSFRow sourceRow = sheet.getRow(i + 1);
            targetRow.copyRowFrom(sourceRow, new CellCopyPolicy());
        }
        //删除最后一行
        sheet.removeRow(sheet.getRow(sheet.getLastRowNum()));
        updateCellRangeAddress();
    }

    /**
     * 插入列
     * 在第sheetIndex个sheet中的第columnIndex列之前插入insertNum列
     * @param sheetIndex    sheetIndex
     * @param columnIndex   columnIndex
     * @param insertNum insertNum
     */
    public void insertColumnsBefore(int sheetIndex, int columnIndex, int insertNum) {
        if (insertNum <= 0) return;
        //TO DO 插入列
        XSSFSheet sheet = this.getSheetAt(sheetIndex);
        if (null == sheet) return ;
        List<int[]> mergeList = new ArrayList<>();
        int maxColumnIndex = -1;    //查找最大列
        for (int i = sheet.getLastRowNum(); i >= 0; i--) {  //从后往前遍历所有行
            XSSFRow row = sheet.getRow(i);
            if (row == null) {  //行为空，则不作任何操作
                continue;
            }
            if (row.getLastCellNum() > maxColumnIndex) {
                maxColumnIndex = row.getLastCellNum();
            }
            for (int j = row.getLastCellNum(); j >= columnIndex; j--) {   //从后往前遍历列之后的所有单元格
                XSSFCell sourceCell = row.getCell(j);
                //创建新的单元格
                XSSFCell targetCell = row.createCell(j + insertNum);
                if (sourceCell == null) { //单元格为空，则不作任何操作
                    continue;
                }
                targetCell.copyCellFrom(sourceCell, new CellCopyPolicy());  //复制单元格
                CellRangeAddress cellAddresses = getCellRangeAddress(sourceCell);   //获取原单元是否为合并单元格
                if (null == cellAddresses) continue;
                int firstRow = cellAddresses.getFirstRow();
                int firstColumn = cellAddresses.getFirstColumn();
                int lastRow = cellAddresses.getLastRow();
                int lastColumn = cellAddresses.getLastColumn();
                if (sourceCell.getRowIndex() != lastRow
                        || sourceCell.getColumnIndex() != lastColumn) continue;
                int regionIndex = cellRangeAddressIntegerMap.get(cellAddresses);
                sheet.removeMergedRegion(cellRangeAddressIntegerMap.get(cellAddresses));
                updateCellRangeAddress();
                if (firstColumn >= columnIndex) {
                    mergeList.add(new int[]{firstRow, lastRow, firstColumn + insertNum, lastColumn + insertNum});
                } else {
                    mergeList.add(new int[]{firstRow, lastRow, firstColumn, lastColumn + insertNum});
                }
            }
            for (int k = columnIndex + insertNum - 1; k >= columnIndex; k--) {
                XSSFCell xssfCell = row.getCell(k);
                if (xssfCell == null) {
                    xssfCell = row.createCell(k);
                }
                CellRangeAddress cellAddresses = getCellRangeAddress(xssfCell);   //获取原单元是否为合并单元格
                if (null == cellAddresses) {
                    row.createCell(k);
                }
            }
        }
        for (int[] merge : mergeList) {
            int firstRow = merge[0];
            int lastRow = merge[1];
            int firstColumn = merge[2];
            int lastColumn = merge[3];
            if (firstColumn < columnIndex && lastColumn >= columnIndex) {
                for (int i = firstRow; i <= lastRow; i++) {
                    XSSFRow row = sheet.getRow(i);
                    if (row == null) {
                        continue;
                    }
                    XSSFCell sourceCell = row.getCell(columnIndex - 1);
                    if (sourceCell == null) {
                        continue;
                    }
                    for (int j = columnIndex; j <= columnIndex + insertNum - 1; j++) {
                        XSSFCell targetCell = row.getCell(j);
                        if (targetCell == null) {
                            targetCell = row.createCell(j);
                        }
                        targetCell.copyCellFrom(sourceCell, new CellCopyPolicy());
                    }
                }
            }
            addMergedRegion(this.getSheetIndex(sheet), firstRow, lastRow, firstColumn, lastColumn);
        }
        //设置列宽
        if (maxColumnIndex > 0) {
            for (int j = maxColumnIndex; j >= columnIndex; j--) {
                int width = sheet.getColumnWidth(j);
                sheet.setColumnWidth(j + insertNum, width);
            }
        }
        for (int j = columnIndex + insertNum - 1; j >= columnIndex; j--) {
            sheet.setColumnWidth(j, DEFAULT_COLUMN_WIDTH);
        }
    }

    /**
     * 插入列
     * 在第sheetIndex个sheet中的第columnIndex列之后插入insertNum列
     * @param sheetIndex    sheetIndex
     * @param columnIndex   columnIndex
     * @param insertNum insertNum
     */
    public void insertColumnsAfter(int sheetIndex, int columnIndex, int insertNum) {
        if (insertNum <= 0) return;
        //TO DO 插入列
        XSSFSheet sheet = this.getSheetAt(sheetIndex);
        if (null == sheet) return ;
        int maxColumnIndex = -1;    //查找最大列
        List<int[]> mergeList = new ArrayList<>();
        for (int i = sheet.getLastRowNum(); i >= 0; i--) {  //从后往前遍历所有行
            XSSFRow row = sheet.getRow(i);
            if (row == null) {  //行为空，则不作任何操作
                continue;
            }
            if (row.getLastCellNum() > maxColumnIndex) {
                maxColumnIndex = row.getLastCellNum();
            }
            for (int j = row.getLastCellNum(); j > columnIndex; j--) {   //从后往前遍历列之后的所有单元格
                XSSFCell sourceCell = row.getCell(j);
                //创建新的单元格
                XSSFCell targetCell = row.createCell(j + insertNum);
                if (sourceCell == null) { //单元格为空，则不作任何操作
                    continue;
                }
                targetCell.copyCellFrom(sourceCell, new CellCopyPolicy());  //复制单元格
                CellRangeAddress cellAddresses = getCellRangeAddress(sourceCell);   //获取原单元是否为合并单元格
                if (null == cellAddresses) continue;
                int firstRow = cellAddresses.getFirstRow();
                int firstColumn = cellAddresses.getFirstColumn();
                int lastRow = cellAddresses.getLastRow();
                int lastColumn = cellAddresses.getLastColumn();
                if (sourceCell.getRowIndex() != lastRow
                        || sourceCell.getColumnIndex() != lastColumn) continue;
                int regionIndex = cellRangeAddressIntegerMap.get(cellAddresses);
                sheet.removeMergedRegion(cellRangeAddressIntegerMap.get(cellAddresses));
                updateCellRangeAddress();
                if (firstColumn > columnIndex) {
                    mergeList.add(new int[]{firstRow, lastRow, firstColumn + insertNum, lastColumn + insertNum});
                } else {
                    mergeList.add(new int[]{firstRow, lastRow, firstColumn, lastColumn + insertNum});
                }
            }
            for (int k = columnIndex + insertNum; k > columnIndex; k--) {
                XSSFCell xssfCell = row.getCell(k);
                if (xssfCell == null) {
                    xssfCell = row.createCell(k);
                }
                CellRangeAddress cellAddresses = getCellRangeAddress(xssfCell);   //获取原单元是否为合并单元格
                if (null == cellAddresses) {
                    row.createCell(k);
                }
            }
        }
        for (int[] merge : mergeList) {
            int firstRow = merge[0];
            int lastRow = merge[1];
            int firstColumn = merge[2];
            int lastColumn = merge[3];
            if (firstColumn <= columnIndex && lastColumn > columnIndex) {
                for (int i = firstRow; i <= lastRow; i++) {
                    XSSFRow row = sheet.getRow(i);
                    if (row == null) {
                        continue;
                    }
                    XSSFCell sourceCell = row.getCell(columnIndex);
                    if (sourceCell == null) {
                        continue;
                    }
                    for (int j = columnIndex + 1; j <= columnIndex + insertNum; j++) {
                        XSSFCell targetCell = row.getCell(j);
                        if (targetCell == null) {
                            targetCell = row.createCell(j);
                        }
                        targetCell.copyCellFrom(sourceCell, new CellCopyPolicy());
                    }
                }
            }
            addMergedRegion(this.getSheetIndex(sheet), firstRow, lastRow, firstColumn, lastColumn);
        }
        //设置列宽
        if (maxColumnIndex > 0) {
            for (int j = maxColumnIndex; j > columnIndex; j--) {
                int width = sheet.getColumnWidth(j);
                sheet.setColumnWidth(j + insertNum, width);
            }
        }
        for (int j = columnIndex + insertNum; j > columnIndex; j--) {
            sheet.setColumnWidth(j, DEFAULT_COLUMN_WIDTH);
        }
    }

    /**
     * 复制列
     * @param sheetIndex    sheetIndex
     * @param targetColumnIndex 目标列
     * @param sourceColumnIndex 源列
     */
    public void copyColumn(int sheetIndex, int targetColumnIndex, int sourceColumnIndex) {
        copyColumn(sheetIndex, targetColumnIndex, sheetIndex, sourceColumnIndex);
    }

    /**
     * 复制列
     * @param targetSheetIndex  目标工作表位置
     * @param targetColumnIndex 目标列位置
     * @param sourceSheetIndex  源工作表位置
     * @param sourceColumnIndex 源列位置
     */
    public void copyColumn(int targetSheetIndex, int targetColumnIndex, int sourceSheetIndex, int sourceColumnIndex) {
        XSSFSheet targetSheet = this.getSheetAt(targetSheetIndex);
        if (targetSheet == null) {
            throw new RenderException("目标工作表不存在！");
        }
        XSSFSheet sourceSheet = this.getSheetAt(sourceSheetIndex);
        if (sourceSheet == null) {
            throw new RenderException("源工作表不存在！");
        }
        int lastRowNum = sourceSheet.getLastRowNum();
        for (int i = 0; i <= lastRowNum; i++) {  //从第一行开始遍历列
            XSSFRow sourceRow = sourceSheet.getRow(i);
            if (sourceRow == null) {
                sourceRow = sourceSheet.createRow(i);
            }
            XSSFRow targetRow = targetSheet.getRow(i);
            if (targetRow == null) {
                targetRow = targetSheet.createRow(i);
            }
            XSSFCell targetCell = targetRow.getCell(targetColumnIndex);
            XSSFCell sourceCell = sourceRow.getCell(sourceColumnIndex);
            if (sourceCell == null && targetCell == null) {
                continue;
            }
            if (targetCell == null) {
                targetCell = targetRow.createCell(targetColumnIndex);
            }
            if (sourceCell == null) {
                sourceCell = sourceRow.createCell(sourceColumnIndex);
            }
            this.copyCell(targetCell, sourceCell);
        }
    }

    /**
     * 移除列
     * @param sheetIndex    工作表位置
     * @param columnIndex   列位置
     */
    public void removeColumn(int sheetIndex, int columnIndex) {
        XSSFSheet sheet = this.getSheetAt(sheetIndex);
        if (sheet == null) {
            throw new RenderException("工作表不存在！");
        }
        XSSFSheet tempSheet = this.createSheet();
        int tempSheetIndex = this.getSheetIndex(tempSheet);
        int maxColumnIndex = -1;
        for (int i = 0 ; i <= sheet.getLastRowNum(); i++) {
            XSSFRow row = sheet.getRow(i);
            if (row == null) continue;
            if (row.getLastCellNum() > maxColumnIndex) {
                maxColumnIndex = row.getLastCellNum();
            }
        }
        //移动
        for (int i = columnIndex; i <= maxColumnIndex; i++) {
            copyColumn(tempSheetIndex, 0, sheetIndex, i + 1);
            for (int j = 0; j <= sheet.getLastRowNum(); j++) {
                XSSFCell tempCell = getCell(sheetIndex, j, i + 1);
                if (isMergedRegionBegin(tempCell)) {
                    removeMergedRegion(tempCell);
                } else if ( i == columnIndex && isMergedRegion(tempCell)) {
                    CellRangeAddress cellAddresses = this.getCellRangeAddress(tempCell);
                    int firstRow = cellAddresses.getFirstRow();
                    int lastRow = cellAddresses.getLastRow();
                    int firstColumn = cellAddresses.getFirstColumn();
                    int lastColumn = cellAddresses.getLastColumn() - 1;
                    this.removeMergedRegion(tempCell);
                    this.addMergedRegion(sheetIndex, firstRow, lastRow, firstColumn, lastColumn);
                }
            }
            copyColumn(sheetIndex, i, tempSheetIndex,0);
        }
        this.removeSheetAt(tempSheetIndex);
    }

    /**
     * 复制单元格
     * @param sheetIndex    工作表
     * @param targetRowIndex    目标行
     * @param targetColumnIndex 目标列
     * @param sourceRowIndex    源行
     * @param sourceColumnIndex 源列
     */
    public void copyCell(int sheetIndex, int targetRowIndex, int targetColumnIndex, int sourceRowIndex, int sourceColumnIndex) {
        copyCell(sheetIndex, targetRowIndex, targetColumnIndex, sheetIndex, sourceRowIndex, sourceRowIndex);
    }

    /**
     * 复制单元格
     * @param targetSheetIndex  目标工作表
     * @param targetRowIndex    目标行
     * @param targetColumnIndex 目标列
     * @param sourceSheetIndex  源工作表
     * @param sourceRowIndex    源行
     * @param sourceColumnIndex 源列
     */
    public void copyCell(int targetSheetIndex, int targetRowIndex, int targetColumnIndex, int sourceSheetIndex, int sourceRowIndex, int sourceColumnIndex) {
        XSSFCell targetCell = getCell(targetSheetIndex, targetRowIndex, targetColumnIndex);
        XSSFCell sourceCell = getCell(sourceSheetIndex, sourceRowIndex, sourceColumnIndex);
        copyCell(targetCell, sourceCell);
    }

    /**
     * 复制单元格
     * @param targetCell    目标单元格
     * @param sourceCell    源单元格
     */
    public void copyCell(XSSFCell targetCell, XSSFCell sourceCell) {
        if (targetCell == null) {
            throw new RenderException("目标单元格为null！");
        }
        if (sourceCell == null) {
            throw new RenderException("源单元格为null！");
        }
        if (isMergedRegionBegin(targetCell)) {  //合并单元格起点，则拆分
            this.removeMergedRegion(targetCell);
        } else if (isMergedRegion(targetCell)) {    //合并单元格，则不作操作
            return ;
        }
        targetCell.copyCellFrom(sourceCell, new CellCopyPolicy());
        if (isMergedRegionBegin(sourceCell)) {   //源单元格为合并单元格起点，则目标单元格需要合并单元格
            CellRangeAddress cellAddresses = this.getCellRangeAddress(sourceCell);
            int firstRow = targetCell.getRowIndex();
            int lastRow = targetCell.getRowIndex() + (cellAddresses.getLastRow() - cellAddresses.getFirstRow());
            int firstColumn = targetCell.getColumnIndex();
            int lastColumn = targetCell.getColumnIndex() + (cellAddresses.getLastColumn() - cellAddresses.getFirstColumn());

            XSSFSheet sourceSheet = sourceCell.getSheet();
            XSSFSheet targetSheet = targetCell.getSheet();
            for (int i = cellAddresses.getFirstRow(); i <= cellAddresses.getLastRow(); i++) {
                XSSFRow sourceRow = sourceSheet.getRow(i);
                if (sourceRow == null) continue;
                int targetRowIndex = targetCell.getRowIndex() + (i - cellAddresses.getFirstRow());
                XSSFRow targetRow = targetSheet.getRow(targetRowIndex);
                if (targetRow == null) {
                    targetRow = targetSheet.createRow(targetRowIndex);
                }
                for (int j = cellAddresses.getFirstColumn(); j <= cellAddresses.getLastColumn(); j++) {
                    XSSFCell tempSourceCell = sourceRow.getCell(j);
                    if (tempSourceCell == null) continue;
                    int tempTargetCellIndex = targetCell.getColumnIndex() + (j - cellAddresses.getFirstColumn());
                    XSSFCell tempTargetCell = targetRow.createCell(tempTargetCellIndex);
                    tempTargetCell.copyCellFrom(tempSourceCell, new CellCopyPolicy());
                }
            }

            this.addMergedRegion(this.getSheetIndex(targetCell.getSheet()), firstRow, lastRow, firstColumn, lastColumn);

        } else if (isMergedRegion(sourceCell)) {
            targetCell.getRow().createCell(targetCell.getColumnIndex());
        }
    }

    /**
     * 获取单元格对象
     * @param sheetIndex    工作表位置
     * @param rowIndex  行位置
     * @param columnIndex   列位置
     * @return
     */
    public XSSFCell getCell(int sheetIndex, int rowIndex, int columnIndex) {
        XSSFSheet sheet = this.getSheetAt(sheetIndex);
        if (sheet == null) {
            throw new RenderException("工作表sheet不存在！");
        }
        XSSFRow row = sheet.getRow(rowIndex);
        if (row == null) {
            row = sheet.createRow(rowIndex);
        }
        XSSFCell cell = row.getCell(columnIndex);
        if (cell == null) {
            cell = row.createCell(columnIndex);
        }
        return cell;
    }

}
