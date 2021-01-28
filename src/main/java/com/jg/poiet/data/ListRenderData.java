package com.jg.poiet.data;

import com.jg.poiet.data.style.Style;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

/**
 * 列数据
 */
public class ListRenderData implements RenderData {

    /**
     * 列数据
     */
    private List<CellRenderData> cellDatas;

    /**
     * 列样式
     */
    private Style style;

    /**
     * 是否扩展行或列（默认扩展）
     */
    private boolean extension = true;

    /**
     * 列表方向（默认垂直）
     */
    private DIRECTION direction = DIRECTION.VERTICAL;

    /**
     * 列表方向枚举
     */
    public enum DIRECTION {
        HORIZONTAL,VERTICAL
    }

    public ListRenderData() {}

    public ListRenderData(List<CellRenderData> cellDatas) {
        this.cellDatas = cellDatas;
    }

    public ListRenderData(List<TextRenderData> cellData, Style style) {
        this.cellDatas = new ArrayList<>();
        if (null != cellData) {
            for (TextRenderData data : cellData) {
                this.cellDatas.add(new CellRenderData(data));
            }
        }
        this.style = style;
    }

    public static ListRenderData build(String... cellStr) {
        List<TextRenderData> cellDatas = new ArrayList<>();
        if (null != cellStr) {
            for (String col : cellStr) {
                cellDatas.add(new TextRenderData(col));
            }
        }
        return new ListRenderData(cellDatas, null);
    }

    public static ListRenderData build(TextRenderData... cellData) {
        return new ListRenderData(null == cellData ? null : Arrays.asList(cellData), null);
    }

    public static ListRenderData build(CellRenderData... cellData) {
        return new ListRenderData(null == cellData ? null : Arrays.asList(cellData));
    }

    public ListRenderData buildStyle(Style style) {
        this.style = style;
        return this;
    }

    public ListRenderData buildExtension(boolean extension) {
        this.extension = extension;
        return this;
    }

    public ListRenderData buildDirection(DIRECTION direction) {
        this.direction = direction;
        return this;
    }

    public int size() {
        return null == cellDatas ? 0 : cellDatas.size();
    }

    public List<CellRenderData> getCellDatas() {
        return cellDatas;
    }

    public void setCellDatas(List<CellRenderData> cellDatas) {
        this.cellDatas = cellDatas;
    }

    public Style getStyle() {
        return style;
    }

    public ListRenderData setStyle(Style style) {
        this.style = style;
        return this;
    }

    public boolean isEmpty() {
        return cellDatas == null || cellDatas.size() <= 0;
    }

    public boolean isNotEmpty() {
        return !isEmpty();
    }

    public boolean isExtension() {
        return extension;
    }

    public void setExtension(boolean extension) {
        this.extension = extension;
    }

    public DIRECTION getDirection() {
        return direction;
    }

    public void setDirection(DIRECTION direction) {
        this.direction = direction;
    }
}
