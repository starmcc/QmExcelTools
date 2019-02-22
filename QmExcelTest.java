package com.qm.code.util.poi;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.IndexedColors;

import java.util.ArrayList;
import java.util.Calendar;
import java.util.List;

/**
 * Copyright © 2018浅梦工作室. All rights reserved.
 *
 * @author 浅梦
 * @date 2018/12/6 11:17
 * @Description: 测试用例
 */
public class QmExcelTest {
    /**
     * main
     */
    public static void main (String[] args){
        // 设置表格格式
        QmExcelFormat qmExcelFormat = new QmExcelFormat();
        qmExcelFormat.setBackgroundColor(IndexedColors.WHITE.getIndex());
        qmExcelFormat.setBold(false);
        qmExcelFormat.setContentHeight((short)500);
        qmExcelFormat.setFontColor(IndexedColors.BLACK.getIndex());
        qmExcelFormat.setFontSize((short)12);
        qmExcelFormat.setFontName("宋体");
        // 创建表格
        QmExcelSheet qmExcelSheet = new QmExcelSheet();
        qmExcelSheet.setTitle("test");
        qmExcelSheet.setSheetName("sheet");
        // 模拟数据插入
        String[] filedNames = new String[]{"字段1","字段2"};
        List<Object[]> contentList = new ArrayList<>();
        for (int i = 0; i < 100000;i++){
            Object[] obj = new Object[2];
            obj[0] = "第" + (i + 3) + "行0列";
            obj[1] = "第" + (i + 3) + "行1列";
            contentList.add(obj);
        }
        qmExcelSheet.setFiledNames(filedNames);
        qmExcelSheet.setContentList(contentList);
        QmExcelContainer qmExcelContainer = new QmExcelContainer(qmExcelFormat);
        // 写入Excel
        HSSFWorkbook book = qmExcelContainer.inputExcel(qmExcelSheet);
        // 输出Excel
        qmExcelContainer.outputSystem("C:\\Users\\Administrator\\Desktop\\test\\"+String.valueOf(Calendar.getInstance().getTimeInMillis()).concat(".xlsx")
                ,book);

    }
}
