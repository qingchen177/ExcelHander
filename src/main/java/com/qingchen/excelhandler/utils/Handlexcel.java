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
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * @author qingchen
 * @date 1/11/2022 下午 1:57
 */

public class Handlexcel {
    public static Map<String, Integer> result = new HashMap<>();

    public static Map<String, Integer> handler() throws IOException {
        List<File> fileList = ReadFiles.readFile("D:\\excel数据处理程序\\表格");

        System.out.println("\t读取文件路径：D:\\excel数据处理程序\\表格");
        System.out.println("\t读取到的EXCEL文件数量为:" + fileList.size());

        System.out.println("\t开始执行...");

        for (File file : fileList) {
            System.out.println("\t当前文件名为：" + file.getName());
            //通过输入流传入excel文件
            FileInputStream in = new FileInputStream(file);

            //将输入流传入Workbook对象并完成解析
            Workbook workbook = new XSSFWorkbook(in);

            //按照名称或下标获取工作簿
            Sheet sheet = workbook.getSheet("表2 动物样线物种记录表");
            int lastRowNum = sheet.getLastRowNum();
            System.out.println("\t当前表行数：" + lastRowNum);

            for (int i = 1; i < lastRowNum; i++) {

                Row row = sheet.getRow(i);
                if (row == null) {
                    System.out.println("\t有效行数为：" + i + 1);
                    break;
                }

                Cell cell1 = row.getCell(3);
                if (cell1 == null || cell1.getStringCellValue() == null || StringUtils.isEmpty(cell1.getStringCellValue())) {
                    System.out.println("\t有效行数为：" + i + 1);
                    break;
                }
                Cell cell2 = row.getCell(4);

                String species = cell1.getStringCellValue();
                Integer number = (int) cell2.getNumericCellValue();
                if (result.containsKey(species)) {
                    result.put(species, result.get(species) + number);
                } else {
                    result.put(species, number);
                }

            }
            System.out.println("\t当前统计结果为：" + result);
        }
        return result;
    }
}
