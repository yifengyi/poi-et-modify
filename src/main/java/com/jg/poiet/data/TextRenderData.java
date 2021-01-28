package com.jg.poiet.data;

import com.jg.poiet.data.style.Style;

/**
 * 文本数据
 *
 */
public class TextRenderData implements RenderData {

    public enum DataType {
        String, Double
    }

    protected Style style;

    /**
     * \n 表示换行
     */
    protected String text;

    /**
     * 数据写入类型（默认字符串）
     */
    protected DataType dataType = DataType.String;

    public TextRenderData() {}

    public TextRenderData(String text) {
        this.text = text;
    }

    public TextRenderData(String text, Style style) {
        this.style = style;
        this.text = text;
    }

    public TextRenderData buildDataType(DataType dataType) {
        this.dataType = dataType;
        return this;
    }

    public Style getStyle() {
        return style;
    }

    public TextRenderData setStyle(Style style) {
        this.style = style;
        return this;
    }

    public String getText() {
        return text;
    }

    public TextRenderData setText(String text) {
        this.text = text;
        return this;
    }

    public DataType getDataType() {
        return dataType;
    }

    public void setDataType(DataType dataType) {
        this.dataType = dataType;
    }
}
