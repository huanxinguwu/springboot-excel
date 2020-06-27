package com.example.mybatistest.excel;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;

import java.io.FileOutputStream;
import java.io.IOException;

/**
 * 将大数据量的数据写入excel2007版，然后导出
 * 相比较2003版，2007版后缀是xlsx，而且支持的数据量无上限，但是速度会慢一些，更消耗内存一些
 */
public class BigData07Test {

    @Test
    public  void bigWrite07() throws IOException {
        //创建workbook
        Workbook workbook=new XSSFWorkbook();
        //创建sheet
        Sheet sheet=workbook.createSheet("07版导出大数据量");
        //创建一个10万行，10列的excel表格
        for (int rowNum=0;rowNum<100000;rowNum++){
            Row row=sheet.createRow(rowNum);
            for (int cellNum=0;cellNum<10;cellNum++){
               Cell cell= row.createCell(cellNum);
               cell.setCellValue(rowNum+"--"+cellNum);
            }
        }
        FileOutputStream fileOutputStream = new FileOutputStream("D:/07版大数据量.xlsx");
        workbook.write(fileOutputStream);
        fileOutputStream.close();
    }

}
