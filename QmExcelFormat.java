package com.qm.code.util.poi;

/**
 * Copyright © 2018浅梦工作室. All rights reserved.
 *
 * @author 浅梦
 * @date 2018/12/6 11:51
 * @Description: QmExcel格式配置类
 */
public class QmExcelFormat {
    private short contentHeight; //对应内容的高度
    private short fontSize; //字体大小
    private short fontColor; //字体颜色 HSSFFont.COLOR_RED
    private String fontName; //字体样式 宋体
    private boolean bold; //字体粗细 true为粗体
    private short backgroundColor; //设置背景颜色

    public short getContentHeight() {
        return contentHeight;
    }

    /**
     * @param contentHeight 对应内容的高度
     */
    public void setContentHeight(short contentHeight) {
        this.contentHeight = contentHeight;
    }

    public short getFontSize() {
        return fontSize;
    }

    /**
     * @param fontSize 字体大小
     */
    public void setFontSize(short fontSize) {
        this.fontSize = fontSize;
    }

    public short getFontColor() {
        return fontColor;
    }

    /**
     * @param fontColor 字体颜色 HSSFFont.COLOR_RED
     */
    public void setFontColor(short fontColor) {
        this.fontColor = fontColor;
    }

    public String getFontName() {
        return fontName;
    }

    /**
     * @param fontName 字体样式 宋体
     */
    public void setFontName(String fontName) {
        this.fontName = fontName;
    }

    public boolean isBold() {
        return bold;
    }

    /**
     * @param bold 字体粗细 true为粗体
     */
    public void setBold(boolean bold) {
        this.bold = bold;
    }

    public short getBackgroundColor() {
        return backgroundColor;
    }

    /**
     * @param backgroundColor 设置背景颜色
     */
    public void setBackgroundColor(short backgroundColor) {
        this.backgroundColor = backgroundColor;
    }
}
