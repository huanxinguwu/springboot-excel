package com.example.mybatistest.excel;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.Test;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

/**
 * 读取excel2003中的数据
 */
public class Excel03Read {

    @Test
    public void read03() throws IOException {
        //读取excel文件
        FileInputStream fileInputStream = new FileInputStream("D:/大数据量.xls");
        //创建workbook
        Workbook workbook=new HSSFWorkbook(fileInputStream);
        //得到sheet
        Sheet sheet=workbook.getSheetAt(0);
        //得到行
        Row row=sheet.getRow(0);
        //得到单元格
        Cell cell=row.getCell(1);
        //得到值
        System.out.println(cell.getStringCellValue());
        //关闭读取流
      fileInputStream.close();


    }




}
