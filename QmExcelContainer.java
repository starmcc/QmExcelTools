package com.qm.code.util.poi;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.formula.functions.T;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;

import javax.servlet.http.HttpServletResponse;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.List;

/**
 * Copyright © 2018浅梦工作室. All rights reserved.
 *
 * @author 浅梦
 * @date 2018/12/6 11:01
 * @Description: Excel容器
 */
public class QmExcelContainer {

    private QmExcelFormat qmExcelFormat;

    /**
     * 实例化一个Excel容器
     *
     * @param qmExcelFormat QmExcel格式实体
     */
    public QmExcelContainer(QmExcelFormat qmExcelFormat) {
        this.qmExcelFormat = qmExcelFormat;
    }

    /**
     * 导出excel文件到指定系统目录文件
     *
     * @param fileName 文件路径全名
     * @param book     要输出的Excel对象
     */
    public void outputSystem(String fileName, HSSFWorkbook book) {
        FileOutputStream fileOut = null;
        try {
            fileOut = new FileOutputStream(fileName);
            book.write(fileOut);
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            try {
                if (fileOut != null) {
                    fileOut.close();
                }
                if (book != null) {
                    book.close();
                }
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }

    /**
     * 导出excel文件到response输出
     *
     * @param response HttpServletResponse
     */
    public void outputResponse(HttpServletResponse response, HSSFWorkbook book) {
        // 生成excel文件
        String fileName = String.valueOf(Calendar.getInstance().getTimeInMillis()).concat(".xls");
        // 清空response
        response.reset();
        response.addHeader("Content-Disposition", "attachment;filename=" + new String(fileName.getBytes()));
        response.setContentType("application/vnd.ms-excel;charset=utf-8");
        OutputStream os = null;
        try {
            os = response.getOutputStream();
            book.write(os);
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            try {
                if (os != null) {
                    os.close();
                }
                if (book != null) {
                    book.close();
                }
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }


    /**
     * 写入Excel数据
     *
     * @param qmExcelSheet 表格数据
     */
    public HSSFWorkbook inputExcel(QmExcelSheet qmExcelSheet) {
        List<Object[]> concatLis = qmExcelSheet.getContentList();
        String sheetName = qmExcelSheet.getSheetName();
        String[] filedNames = qmExcelSheet.getFiledNames();
        String title = qmExcelSheet.getTitle();
        int endNum = 10000;
        List<List<Object[]>> concatLisSplit = splitList(concatLis, endNum);
        // 创建一个Excel表格
        HSSFWorkbook book = new HSSFWorkbook();
        for (int i = 0; i < concatLisSplit.size(); i++) {
            // 创建一个单元表
            HSSFSheet sheet = book.createSheet(sheetName + (i + 1));
            // 合并单元格
            sheet.addMergedRegion(new CellRangeAddress(
                    0, 0, 0, filedNames.length - 1));
            //创建第一行
            HSSFRow titleRow = sheet.createRow(0);
            //设置第一行单元格高度
            titleRow.setHeight((short) 1200);
            //创建第一行第一列单元格
            HSSFCell titleCell = titleRow.createCell(0);
            //设置单元格的值
            titleCell.setCellValue(title);
            //3.单元格使用样式，设置第一行第一列单元格样式
            titleCell.setCellStyle(getTitleStyle(book));

            // ========创建字段行========
            // 获取样式
            HSSFCellStyle filedCellStyle = getContentStyle(book);
            //创建第1行 因为标题占了第0行
            HSSFRow filedRow = sheet.createRow(1);
            //设置字段行单元格高度
            filedRow.setHeight((short) 500);
            for (int f = 0; f < filedNames.length; f++) {
                //创建字段行第i列单元格
                HSSFCell filedCell = filedRow.createCell(f);
                filedCell.setCellValue(filedNames[f]);
                filedCell.setCellStyle(filedCellStyle);
                //自适应宽度
                sheet.autoSizeColumn(f);
                sheet.setColumnWidth(f, sheet.getColumnWidth(f) * 18 / 10);
            }
            System.out.println("创建了[" + sheetName + i + "]表");
            // ===========创建内容==========
            HSSFCellStyle contentCellStyle = getContentStyle(book);
            List<Object[]> contentList = concatLisSplit.get(i);
            for (int c = 0; c < contentList.size(); c++) {
                //创建第i + 2行 因为标题和字段共占了2行
                HSSFRow contentRow = sheet.createRow(c + 2);
                //设置字段行单元格高度
                contentRow.setHeight(qmExcelFormat.getContentHeight());
                for (int j = 0; j < contentList.get(c).length; j++) {
                    HSSFCell contentCell = contentRow.createCell(j);
                    contentCell.setCellValue(contentToString(contentList.get(c)[j]));
                    contentCell.setCellStyle(contentCellStyle);
                }
            }
        }
        return book;
    }

    /**
     * 分解list
     *
     * @param list      元数据
     * @param groupSize 分解大小
     * @param <T>       泛型
     * @return
     */
    private <T> List<List<T>> splitList(List<T> list, int groupSize) {
        int length = list.size();
        // 计算可以分成多少组
        int num = (length + groupSize - 1) / groupSize; // TODO
        List<List<T>> newList = new ArrayList<>(num);
        for (int i = 0; i < num; i++) {
            // 开始位置
            int fromIndex = i * groupSize;
            // 结束位置
            int toIndex = (i + 1) * groupSize < length ? (i + 1) * groupSize : length;
            newList.add(list.subList(fromIndex, toIndex));
        }
        return newList;
    }

    /**
     * @param
     * @return String
     * @Title: contentToString
     * @Description: 转换任何类型toString
     */
    private String contentToString(Object obj) {
        if (obj == null) {
            return "";
        }
        String result;
        try {
            if (obj instanceof Integer
                    || obj instanceof Long
                    || obj instanceof Double
                    || obj instanceof Float
                    || obj instanceof Boolean) {
                result = String.valueOf(obj);

            } else if (obj instanceof Date) {
                SimpleDateFormat fomt = new SimpleDateFormat("yyyy-MM-dd hh:mm:ss");
                result = fomt.format(obj);
            } else {
                result = obj.toString();
            }
        } catch (Exception e) {
            //e.printStackTrace();
            result = "数据错误";
        }
        return result;
    }


    /**
     * 设置内容样式
     *
     * @param book book
     * @return
     */
    private HSSFCellStyle getContentStyle(HSSFWorkbook book) {
        //改变字体样式，步骤
        HSSFFont hssfFont = book.createFont();
        //设置字体颜色
        hssfFont.setColor(qmExcelFormat.getFontColor());
        //字体粗体显示
        hssfFont.setBold(qmExcelFormat.isBold());
        hssfFont.setFontName(qmExcelFormat.getFontName());
        // 字体大小
        hssfFont.setFontHeightInPoints(qmExcelFormat.getFontSize());
        //设置样式
        HSSFCellStyle cellStyle = book.createCellStyle();
        cellStyle.setFont(hssfFont);
        //设置单元格背景色
        cellStyle.setFillForegroundColor(qmExcelFormat.getBackgroundColor());
        cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        //设置居中
        cellStyle.setAlignment(HorizontalAlignment.CENTER);//水平居中
        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);//垂直居中
        //设置边框
        cellStyle.setBorderBottom(BorderStyle.THIN); //下边框
        cellStyle.setBorderLeft(BorderStyle.THIN);//左边框
        cellStyle.setBorderTop(BorderStyle.THIN);//上边框
        cellStyle.setBorderRight(BorderStyle.THIN);//右边框
        return cellStyle;
    }

    /**
     * 设置字段样式
     *
     * @param book book
     * @return
     */
    private HSSFCellStyle getFildStyle(HSSFWorkbook book) {
        //改变字体样式，步骤
        HSSFFont hssfFont = book.createFont();
        //设置字体颜色
        hssfFont.setColor(IndexedColors.RED.getIndex());
        //字体粗体显示
        hssfFont.setBold(true);
        hssfFont.setFontName(qmExcelFormat.getFontName());
        // 字体大小
        hssfFont.setFontHeightInPoints((short) 14);
        //设置样式
        HSSFCellStyle cellStyle = book.createCellStyle();
        cellStyle.setFont(hssfFont);
        //设置居中
        cellStyle.setAlignment(HorizontalAlignment.CENTER);//水平居中
        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);//垂直居中
        //设置边框
        cellStyle.setBorderBottom(BorderStyle.THIN); //下边框
        cellStyle.setBorderLeft(BorderStyle.THIN);//左边框
        cellStyle.setBorderTop(BorderStyle.THIN);//上边框
        cellStyle.setBorderRight(BorderStyle.THIN);//右边框
        return cellStyle;
    }


    /**
     * 设置标题样式
     *
     * @param book book
     * @return
     */
    private HSSFCellStyle getTitleStyle(HSSFWorkbook book) {
        //改变字体样式，步骤
        HSSFFont hssfFont = book.createFont();
        //设置字体颜色
        hssfFont.setColor(IndexedColors.WHITE1.getIndex());
        //字体粗体显示
        hssfFont.setBold(true);
        hssfFont.setFontName(qmExcelFormat.getFontName());
        // 字体大小
        hssfFont.setFontHeightInPoints((short) 28);
        //设置样式
        HSSFCellStyle cellStyle = book.createCellStyle();
        cellStyle.setFont(hssfFont);

        //设置单元格背景色
        cellStyle.setFillForegroundColor(IndexedColors.BLUE1.getIndex());
        cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        //设置居中
        cellStyle.setAlignment(HorizontalAlignment.CENTER);//水平居中
        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);//垂直居中
        //设置边框
        cellStyle.setBorderBottom(BorderStyle.THIN); //下边框
        cellStyle.setBorderLeft(BorderStyle.THIN);//左边框
        cellStyle.setBorderTop(BorderStyle.THIN);//上边框
        cellStyle.setBorderRight(BorderStyle.THIN);//右边框
        return cellStyle;
    }
}
