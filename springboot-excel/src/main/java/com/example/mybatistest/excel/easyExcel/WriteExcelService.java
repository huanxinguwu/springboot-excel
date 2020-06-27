package com.example.mybatistest.excel.easyExcel;

import com.alibaba.excel.EasyExcel;
import org.junit.Test;

import java.util.ArrayList;
import java.util.List;

/**
 * 往excel表格里写入数据
 */
public class WriteExcelService {
    //将需要输出的数据写入list
    private List<ExcelEntity> data(){
        List<ExcelEntity> list=new ArrayList<ExcelEntity>();
        for (int i=0;i<10;i++){
            ExcelEntity writeExcelEntity=new ExcelEntity();
            writeExcelEntity.setName("李白"+i);
            writeExcelEntity.setAge(10+i);
            writeExcelEntity.setScore(60+i);
            list.add(writeExcelEntity);
        }
        return list;
    }
    //将list中的值写入excel,生成excel
    @Test
    public void simpWrite(){
        //指定输出的excel的路径以及名称
        String fileName="D:/分数统计表.xlsX";
        //指定按照哪个实体写数据，默认写入第一个sheet，并指定sheet名称
        EasyExcel.write(fileName,ExcelEntity.class).sheet("分数统计").doWrite(data());
    }
}
