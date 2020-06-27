package com.example.mybatistest.excel;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

/**
 * 2007版支持写入更多数据，但是就是效率低一点，
 * 但是这里有优化策略，创建workbook时，使用SXSSFWorkbook创建
 * 因为使用SXSSFWorkbook创建workbook的方式会通过产生临时文件来提高效率
 */
public class BIgData07FastWrite {
    @Test
    public void fastWrite() throws IOException {
        //使用SXSSFWorkbook创建workbbok
        Workbook workbook=new SXSSFWorkbook();
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
        FileOutputStream fileOutputStream = new FileOutputStream("D:/07版快速导出大数据量.xlsx");
        workbook.write(fileOutputStream);
        fileOutputStream.close();
        //清除掉workbook产生的临时文件
        ((SXSSFWorkbook)workbook).dispose();
    }


    }

