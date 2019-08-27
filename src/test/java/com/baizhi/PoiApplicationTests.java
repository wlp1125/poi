package com.baizhi;

import org.apache.poi.hssf.usermodel.*;
import org.junit.Test;
import org.junit.runner.RunWith;
import org.springframework.boot.test.context.SpringBootTest;
import org.springframework.test.context.junit4.SpringRunner;

import java.io.File;
import java.io.IOException;
import java.util.Date;

@RunWith(SpringRunner.class)
@SpringBootTest
public class PoiApplicationTests {

    @Test
    public void contextLoads() {
        //创建Excel工作簿对象
        HSSFWorkbook workbook = new HSSFWorkbook();
        // 创建工作表
        HSSFSheet sheet = workbook.createSheet("用户信息");
        // 创建标题行
        HSSFRow row = sheet.createRow(0);
        String[] title = {"编号", "姓名", "出生年月"};
        // 创建单元格对象
        HSSFCell cell = null;
        for (int i = 0; i < title.length; i++) {
            // i标题列索引
            cell = row.createCell(i);
            // 给单元格赋值
            cell.setCellValue(title[i]);
        }
        //处理日期格式    由工作簿对象创建样式对象
        HSSFCellStyle cellStyle = workbook.createCellStyle();
        //由工作簿对象创建日期格式
        HSSFDataFormat dataFormat = workbook.createDataFormat();
        //设置日期格式 由样式对象设置日期格式
        cellStyle.setDataFormat(dataFormat.getFormat("yyyy年MM月dd日"));
        //处理数据行
        for (int i = 1; i < 10; i++) {
            row = sheet.createRow(i);
            row.createCell(0).setCellValue(i);
            row.createCell(1).setCellValue("老谈" + 1);
            //设置日期格式
            cell = row.createCell(2);
            cell.setCellValue(new Date());
            cell.setCellStyle(cellStyle);

        }
        try {
            workbook.write(new File("e:\\用户.xls"));
            workbook.close();
        } catch (IOException e) {
            e.printStackTrace();
        }

    }

}
