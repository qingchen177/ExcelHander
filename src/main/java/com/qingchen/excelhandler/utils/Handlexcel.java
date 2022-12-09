package com.qingchen.excelhandler.utils;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.util.StringUtils;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * @author qingchen
 * @date 1/11/2022 下午 1:57
 */

public class Handlexcel {
    public static Map<String, Integer> startResult = new HashMap<>();
    public static Map<String, Integer> endResult = new HashMap<>();

    public static List<Map<String, Integer>> handler() throws IOException {
        List<File> fileList = ReadFiles.readFile("D:\\excel数据处理程序-Version-2.0\\表格");

        System.out.println("\t读取文件路径：D:\\excel数据处理程序-Version-2.0\\表格");
        System.out.println("\t读取到的EXCEL文件数量为:" + fileList.size());

        System.out.println("\t开始执行...");

        for (File file : fileList) {
            System.out.println("\t当前文件名为：" + file.getName());
            //通过输入流传入excel文件
            FileInputStream in = new FileInputStream(file);

            //将输入流传入Workbook对象并完成解析
            Workbook workbook = new XSSFWorkbook(in);

            //按照名称或下标获取工作簿
            Sheet sheet = workbook.getSheet("表1 动物调查样线信息表");
            int lastRowNum = sheet.getLastRowNum();
            System.out.println("\t当前表行数：" + lastRowNum);

            for (int i = 1; i < lastRowNum; i++) {

                Row row = sheet.getRow(i);
                if (row == null) {
                    System.out.println("\t有效行数为：" + i + 1);
                    break;
                }

                Cell cell1 = row.getCell(13);
                if (cell1 == null || cell1.getStringCellValue() == null || StringUtils.isEmpty(cell1.getStringCellValue())) {
                    System.out.println("\t有效行数为：" + i + 1);
                    break;
                }
                Cell cell2 = row.getCell(23);

                String startValue = cell1.getStringCellValue();
                String endValue = cell2.getStringCellValue();

                startResult.put(startValue, startResult.getOrDefault(startValue,0) + 1 );
                endResult.put(endValue, endResult.getOrDefault(endValue,0) + 1 );

            }
            System.out.println("\t当前统计结果为：" + startResult + "\n" + endResult + "\n");
        }
        List<Map<String, Integer>> result = new ArrayList<>();
        result.add(startResult);
        result.add(endResult);
        return result;
    }
}
