package com.qingchen.excelhandler.utils;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Iterator;
import java.util.List;
import java.util.Map;

/**
 * @author qingchen
 * @date 1/11/2022 下午 3:21
 */

public class Createxcel {
    public static void createExcel(List<Map<String, Integer>> maps) {
        try (Workbook workbook = new XSSFWorkbook();
             FileOutputStream file = new FileOutputStream("D:\\excel数据处理程序-Version-2.0\\干扰类型统计结果.xlsx")) {

            //创建sheet工作簿
            Sheet sheet0 = workbook.createSheet("干扰类型统计表");

            //创建row行表头
            Row row0 = sheet0.createRow(0);

            //创建cell单元格并写入表头数据
            Cell cell0 = row0.createCell(0);
            cell0.setCellValue("起点生境受干扰类型");
            Cell cell1 = row0.createCell(1);
            cell1.setCellValue("类型数量");
            Cell cell2 = row0.createCell(2);
            cell2.setCellValue("终点生境受干扰类型");
            Cell cell3 = row0.createCell(3);
            cell3.setCellValue("类型数量");

            Map<String, Integer> startMap = maps.get(0);
            Map<String, Integer> endMap = maps.get(1);

            Iterator<Map.Entry<String, Integer>> startIterator = startMap.entrySet().iterator();
            Iterator<Map.Entry<String, Integer>> endIterator = endMap.entrySet().iterator();

            int size = Math.max(startMap.size(), endMap.size());

            //写入数据
            for (int i =  1; i <= size; i++){
                Row row = sheet0.createRow(i);

                if (startIterator.hasNext()){
                    Map.Entry<String, Integer> entry = startIterator.next();
                    row.createCell(0).setCellValue(entry.getKey());
                    row.createCell(1).setCellValue(entry.getValue());
                    startIterator.remove();
                }

                if (endIterator.hasNext()){
                    Map.Entry<String, Integer> entry = endIterator.next();
                    row.createCell(2).setCellValue(entry.getKey());
                    row.createCell(3).setCellValue(entry.getValue());
                    endIterator.remove();
                }

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
