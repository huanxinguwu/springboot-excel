package com.example.mybatistest.excel.easyExcel;

import com.alibaba.excel.EasyExcel;
import org.junit.Test;


public class ReadExcelService {

    @Test
    public void simpleRead(){
        //需要读取的数据，在实际开发中应该是先上传excel文件，再解析excel，这里就直接指定本地的文件了
        String fileName="D:/分数统计表.xlsX";
        //这里需要指定读取的文件，用哪个class去读，
        // 读取第几个sheet的数据，默认是第一个sheet
        EasyExcel.read(fileName,ExcelEntity.class,new ReadExcelListener()).sheet().doRead();
    }
}
