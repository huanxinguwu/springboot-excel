package com.example.mybatistest.excel.exportAndImportExcel;
import io.swagger.annotations.Api;
import io.swagger.annotations.ApiOperation;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;
import org.springframework.web.multipart.MultipartHttpServletRequest;

import javax.annotation.Resource;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.File;
import java.io.IOException;
import java.io.OutputStream;
import java.nio.charset.StandardCharsets;
import java.time.LocalDate;
import java.util.ArrayList;
import java.util.List;

@RestController
@Api(tags = "【Excel批量导入导出】")
public class TestController {

    @Resource
    private ExcelTestMapper excelTestMapper;

    @Autowired
    private ExcelTestServiceImpl excelTestService;
//    @ApiOperation("根据ids查询统计")
//    @PostMapping("/count")
//    public void counts(){
//        int[] ids={0,1,2};
//        List<Integer> list=new ArrayList<>();
//        list.add(0);
//        list.add(1);
//        list.add(2);
//        List<String> lists=testMapper.selectCount(list);
//        System.out.println(lists);
//
//    }

    @ApiOperation("【导出excel】")
    @GetMapping("exportExcel")
    public void exportExcel(HttpServletResponse response) throws IOException {
         //String fileName="分数统计表";
         //解决文件名中文乱码
        String fileName = new String(("score" + ".xlsx").getBytes(StandardCharsets.UTF_8), StandardCharsets.UTF_8);
          OutputStream outputStream=null;
//        SXSSFWorkbook wb=excelTestService.exportExcel()
        List<ExcelTest> list=excelTestMapper.selectAll();
        Workbook wb=excelTestService.exportExcel(list);
        response.setContentType("application/vnd.ms-excel");
        response.setHeader("Content-disposition","attachment;filename="+fileName);
       outputStream=response.getOutputStream();
       wb.write(outputStream);
       outputStream.flush();
       outputStream.close();
    }

    //上传excel数据并解析，插入数据库
    @ApiOperation("【上传excel并解析】")
    @PostMapping("/deliverExcel")
    public void deliverExcel(@RequestParam("file") MultipartFile file) throws Exception {

        String originalFilename=file.getOriginalFilename();
       Long size= file.getSize();

        System.out.println("originalFilename是："+originalFilename);
       System.out.println("文件大小是："+size);
        System.out.println("w文件名是"+file.getName());
        System.out.println("resource是"+file.getResource());
        excelTestService.deliverExcel(file);

    }

    //使用mybatis配置文件版通过id查询数据
    @ApiOperation("【注解版查数据】")
    public  ExcelTest selectByid(int id){
        return excelTestMapper.selectById(id);
    }

}
