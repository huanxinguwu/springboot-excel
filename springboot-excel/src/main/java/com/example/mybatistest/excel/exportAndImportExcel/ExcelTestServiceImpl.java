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
    public SXSSFWorkbook exportExcel(List<ExcelTest> list) {
        //查询到的数据
       // List<ExcelTest> list = excelTestMapper.selectAll();
        //开始讲数据封装到excel
        SXSSFWorkbook workbook = new SXSSFWorkbook();
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

    //将解析到的excel数据插入数据库
    public void deliverExcel(MultipartFile multipartFile) throws IOException, InvalidFormatException {
     //String filePath=file.get();
        //读取excel文件
        InputStream fileInputStream = multipartFile.getInputStream();
        System.out.println("得到了文件流");
        int file=fileInputStream.read();
        System.out.println("文件流的长度是"+file );
        String filePath=multipartFile.getOriginalFilename();
        //创建文件输入流
        //FileInputStream fileInputStreams = new FileInputStream(filePath);
        //创建workbook
        XSSFWorkbook workbook=new XSSFWorkbook(fileInputStream);
//        Workbook workbook = WorkbookFactory.create(fileInputStream);
        //得到sheet
        Sheet sheet=workbook.getSheetAt(0);
        //得到行
        Row row=sheet.getRow(0);
        //得到单元格
        Cell cell=row.getCell(0);
        //得到值
        System.out.println(cell.getStringCellValue());
        fileInputStream.close();

    }
}
