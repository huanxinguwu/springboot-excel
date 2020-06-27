package com.example.mybatistest.excel.easyExcel;

import org.springframework.stereotype.Service;

import java.util.List;

@Service
/**
 * 在实际开发中，还是要按照mvc的规范来写数据的持久化，
 * 这里只要是为了方便演示效果，就不写入库的操作了
 */
public class ExccelMapper  {
    //持久化操作
    public void saveData(List<ExcelEntity> list){};
}
