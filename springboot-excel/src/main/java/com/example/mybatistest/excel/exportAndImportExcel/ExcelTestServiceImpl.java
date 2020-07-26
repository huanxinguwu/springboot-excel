package com.example.mybatistest.excel.exportAndImportExcel;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import javax.annotation.Resource;
import java.io.*;
import java.util.List;

@Service
public class ExcelTestServiceImpl {

    @Resource
    private ExcelTestMapper excelTestMapper;

    /**
     * 查询所有数据并封装到excel
     */
    public Workbook exportExcel(List<ExcelTest> list) {
        //查询到的数据
       // List<ExcelTest> list = excelTestMapper.selectAll();
        //开始讲数据封装到excel
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("学生成绩表");
        //在sheet中添加表头第0行
        Row row1 = sheet.createRow(0);
        Cell ceLL11 = row1.createCell(0);
        ceLL11.setCellValue("id");
        //创建第一行第二列
        Cell ceLL12=row1.createCell(1);
        ceLL12.setCellValue("姓名");
        //创建第一行第三列
        Cell ceLL13=row1.createCell(2);
        ceLL13.setCellValue("性别");
        //创建第一行第四列
        Cell ceLL14=row1.createCell(3);
        ceLL14.setCellValue("分数");
        int length=list.size();
        for (int i=0;i<length;i++){
            Row row=sheet.createRow(i+1);
            row.createCell(0).setCellValue(list.get(i).getId());
            row.createCell(1).setCellValue(list.get(i).getUserName());
            row.createCell(2).setCellValue(list.get(i).getGender());
            row.createCell(3).setCellValue(list.get(i).getScore());
        }
        return  workbook;
    }

    //上传excel，并将解析到的excel数据插入数据库
    public void deliverExcel(MultipartFile multipartFile) throws Exception {
     String fileName=multipartFile.getOriginalFilename();
        System.out.println(fileName);
        //获取文件流，读取excel文件
        FileInputStream fileInputStream = (FileInputStream)multipartFile.getInputStream();
        String filePath=multipartFile.getOriginalFilename();
        Workbook workbook = null;
        String[] fileNameArray=fileName.split("\\.");
        //获取后缀名
        String extentionName=fileNameArray[fileNameArray.length-1];
        System.out.println(extentionName);
        if (extentionName.equalsIgnoreCase("xls")) {
            // 创建2003版excel
            workbook = new HSSFWorkbook(fileInputStream);
        } else if (extentionName.equalsIgnoreCase("XLSX")) {
            // 创建2007版excel
            workbook = new XSSFWorkbook(fileInputStream);
        } else {
            throw new Exception("文件不是Excel文件");
        }
        //得到sheet
        Sheet sheet=workbook.getSheet("学生成绩表");
        int rowLength=sheet.getPhysicalNumberOfRows();
        System.out.println("总行数是："+rowLength);
        //因为第一行写的是各个列的中文，例如id，姓名，性别，分数，所以从i 从1 开始循环
        for (int i=1;i<rowLength;i++){
            //得到行
            Row row=sheet.getRow(i);
            //得到第i行第一单元格
            Cell cell00=row.getCell(0);
            //得到值
            int cell1Value=(int)cell00.getNumericCellValue();
            //第二个单元格的值
           Cell cell01=row.getCell(1);
           String cell2Value=cell01.getStringCellValue();
            //第三个单元格值
            Cell cell02=row.getCell(2);
            String cell3Value=cell02.getStringCellValue();
            //第四个单元格的值
           Cell cell03=row.getCell(3);
            int cell4Value=(int)cell03.getNumericCellValue();
            System.out.println(cell1Value+"-"+cell2Value+"-"+cell3Value+"-"+cell4Value);
        }

        fileInputStream.close();

    }
}
