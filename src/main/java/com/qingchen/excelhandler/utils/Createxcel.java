package com.qingchen.excelhandler.utils;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Map;

/**
 * @author qingchen
 * @date 1/11/2022 下午 3:21
 */

public class Createxcel {
    public static void createExcel(Map<String, Integer> map) {
        try (Workbook workbook = new XSSFWorkbook();
             FileOutputStream file = new FileOutputStream("D:\\excel数据处理程序\\物种统计结果.xlsx")) {

            //创建sheet工作簿
            Sheet sheet0 = workbook.createSheet("物种统计结果表");

            int i = 0;
            //创建row行
            Row row0 = sheet0.createRow(i);

            //创建cell单元格并写进内容
            Cell cell0 = row0.createCell(0);
            cell0.setCellValue("物种名称");
            Cell cell1 = row0.createCell(1);
            cell1.setCellValue("实体数量");


            for (Map.Entry<String, Integer> mapSet : map.entrySet()) {
                Row row = sheet0.createRow(++i);
                Cell species = row.createCell(0);
                species.setCellValue(mapSet.getKey());
                Cell number = row.createCell(1);
                number.setCellValue(mapSet.getValue());
            }


            //写入
            System.out.println("正在写入数据...");
            workbook.write(file);

            System.out.println("数据写入完毕!!!");

        } catch (IOException e) {
            e.printStackTrace();
        }

    }
}
