package com.example.mybatistest.excel;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.Test;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

/**
 * 写入大量数据到excel2003，excel2003最多支持65536行数据，多余65536行就报错
 * 相比2007版，excel2003的写入速度会更快，内存消耗也更少，但是缺点就是最多只能写65536行
 */
public class BigData03Write {

    @Test
    public void bigDataWrite() throws IOException {
        //创建workbook
        Workbook workbook = new HSSFWorkbook();
        //创建一个sheet
        Sheet sheet = workbook.createSheet("大数据量");
        //写入一个65536行，10列的数据
        for (int rowNum = 0; rowNum < 65536; rowNum++) {
            Row row=sheet.createRow(rowNum);
            for (int cellNum=0;cellNum<10;cellNum++){
                Cell cell=row.createCell(cellNum);
                cell.setCellValue(rowNum+"--"+cellNum);
            }
        }
        //将生成的excel导出
        FileOutputStream fileOutputStream = new FileOutputStream("D:/大数据量.xls");
        workbook.write(fileOutputStream);
        fileOutputStream.close();
    }
}



