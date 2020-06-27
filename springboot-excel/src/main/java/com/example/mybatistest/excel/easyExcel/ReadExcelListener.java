package com.example.mybatistest.excel.easyExcel;

import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.event.AnalysisEventListener;
import com.alibaba.fastjson.JSON;
import org.springframework.beans.factory.annotation.Autowired;

import java.util.ArrayList;
import java.util.List;

/**
 * 使用easyexcel读取数据需要先建立监听器监听excel
 */
public class ReadExcelListener extends AnalysisEventListener<ExcelEntity> {
    //如果有需要将读取到的数据插入数据库的话，默认是每5条插入数据库
    private static  final int BATCH_SIZE =5;
    List<ExcelEntity> list=new ArrayList<ExcelEntity>();
//
        //这里就不插入数据库里了，直接将读取的数据输出出来作为演示
    /**
     *自动注入dao
     *
     */
    @Autowired
   private ExccelMapper excelMapper;
//   public ReadExcelListener(){
//        excelMapper=new ExccelMapper();
//    }
//    public ReadExcelListener(){
//        this.excelMapper=excelMapper;
//    }
    /**
     * 数据读取的核心方法，读取数据会执行这个方法
     * ExcelEntity：表格对应的实体
     * AnalysisContext：分析上下文
     * @param excelEntity
     * @param analysisContext
     */
    @Override
    public void invoke(ExcelEntity excelEntity, AnalysisContext analysisContext) {
        //将读取到的数据输出，JSON.toJSONString(excelEntity就是通过fastjson将数据转化
        System.out.println(JSON.toJSONString(excelEntity));
        list.add(excelEntity);
        //当读取的数据量达到设定的批量规模后，执行插入数据库方法,
        // 否则将读取的数据都读取出来不插入的话，容易造成内存溢出,下面入库的方法这里就不写了
//        if (list.size()>=BATCH_SIZE){
//            excelMapper.saveData(list);
//            //将插入成功的数据从list清除掉
//            list.clear();
//        }
    }

    /**
     * 所有数据读取完了都会调用的方法
     * @param analysisContext
     */
    @Override
    public void doAfterAllAnalysed(AnalysisContext analysisContext) {
           //在这里也要保存数据，确保最后遗留的数据也存储到数据库
        //excelMapper.saveData(list);
        System.out.println("所有数据解析完成");
    }
}
