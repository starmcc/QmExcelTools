package com.qm.code.util.poi;


import java.util.List;

/**
 * Copyright © 2018浅梦工作室. All rights reserved.
 *
 * @author 浅梦
 * @date 2018/12/6 1:04
 * @Description: Excel的Sheet表格实体类
 */
public class QmExcelSheet {
    private String sheetName;//表名
    private String title; //标题
    private String[] filedNames; //字段名数组
    private List<Object[]> contentList; //对集合的数组对象 该数组必须和字段名数组对应顺序和长度

    public String getSheetName() {
        return sheetName;
    }

    public void setSheetName(String sheetName) {
        this.sheetName = sheetName;
    }

    public String getTitle() {
        return title;
    }

    public void setTitle(String title) {
        this.title = title;
    }

    public String[] getFiledNames() {
        return filedNames;
    }

    public void setFiledNames(String[] filedNames) {
        this.filedNames = filedNames;
    }

    public List<Object[]> getContentList() {
        return contentList;
    }

    public void setContentList(List<Object[]> contentList) {
        this.contentList = contentList;
    }
}
