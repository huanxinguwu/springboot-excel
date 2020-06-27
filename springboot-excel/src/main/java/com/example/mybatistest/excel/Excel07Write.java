package com.example.mybatistest.excel;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;

import java.io.FileOutputStream;

/**
 * 是基于excel2007生成的，与excel2003版区别仅仅在于两点：
 * 1.2003版是将生成的excel用xls后缀，2007版是将生成的excel使用xlsx后缀结尾
 * 2.2003版是用HSSFWorkbook生成工作薄，2007版是使用XSSFWorkbook生成工作簿
 */
public class Excel07Write {
    //使用excel2003  将数据批量导出到指定路径
    @Test
    public void test01() throws Exception {
        //创建一个工作薄
        Workbook workbook=new XSSFWorkbook();
        //创建一个工作表，就是excel表格下方的sheet，并给sheet起名称
        Sheet sheet=workbook.createSheet("excel2007分数统计表");
        //创建第一个行,对应excel的第一行(最起始的行)
        Row row1=sheet.createRow(0);
        //创建一个单元格,就是第一行第一列的单元格
        Cell ceLL11=row1.createCell(0);
        //给ceLL11(第一行第一列的单元格)赋值
        ceLL11.setCellValue("姓名");
        //创建第一行第二列的单元格
        Cell ceLL12=row1.createCell(1);
        //给第一行第二列赋值
        ceLL12.setCellValue("分数");
        //创建第二行
        Row row2=sheet.createRow(1);
        //创建第二行的第一列
        Cell ceLL21=row2.createCell(0);
        //给ceLL11(第一行第一列的单元格)赋值
        ceLL21.setCellValue("李白");
        //创建第二行第二列的单元格
        Cell ceLL22=row2.createCell(1);
        //给第二行第二列赋值
        ceLL22.setCellValue("97");
        //创建第三行
        Row row3=sheet.createRow(2);
        //创建第三行的第一列
        Cell ceLL31=row3.createCell(0);
        //给第三行的第一列赋值
        ceLL31.setCellValue("刘备");
        //给第三行的第二列的单元格
        Cell ceLL32=row3.createCell(1);
        //给第三行第二列赋值
        ceLL32.setCellValue(98);
        //将生成的excel文件输出(IO流) 2003版使用xlsx后缀结尾
        FileOutputStream fileOutputStream = new FileOutputStream("D:/分数统计表.xlsX");
        workbook.write(fileOutputStream);
        //关闭流
        fileOutputStream.close();
        System.out.println("excel生成完毕");
    }




}
