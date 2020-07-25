package com.example.mybatistest.excel.exportAndImportExcel;
import com.example.mybatistest.excel.easyExcel.ExcelEntity;
import org.apache.ibatis.annotations.Insert;
import org.apache.ibatis.annotations.Mapper;
import org.apache.ibatis.annotations.Select;

import java.util.List;

@Mapper
public interface ExcelTestMapper {
      //查询所有数据
        @Select("select * from excel_test")
        List<ExcelTest> selectAll();

        //将解析到的excel数据插入表
        @Insert("insert into excel_test(id,user_name,gender,score) values(#{id},#{userName},#{gender},#{score})")
        int insertExcel();
}
