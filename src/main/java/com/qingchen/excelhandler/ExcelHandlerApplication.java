package com.qingchen.excelhandler;

import com.qingchen.excelhandler.utils.Createxcel;
import com.qingchen.excelhandler.utils.Handlexcel;
import lombok.SneakyThrows;
import lombok.extern.slf4j.Slf4j;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;

import java.net.UnknownHostException;
import java.util.Map;

@Slf4j
@SpringBootApplication
public class ExcelHandlerApplication {

    @SneakyThrows
    public static void main(String[] args) throws UnknownHostException {
        SpringApplication.run(ExcelHandlerApplication.class, args);
        System.out.println("------------------------清尘----------------------------\n" +
                        "应用 'ExcelHandler' 运行成功! \n" +
                        "----------------------------------------------------------\n");
        System.out.println("正在进行数据解析业务...");

        Map<String, Integer> result = Handlexcel.handler();

        System.out.println("数据解析完毕...");
        System.out.println("开始写入数据并生成EXCEL文件...");

        Createxcel.createExcel(result);

        System.out.println("文件已生成，文件路径为：D:\\excel数据处理程序\\物种统计结果.xlsx ");
        System.out.println("程序执行完毕!!!");
        System.out.println("已退出。");
    }

}
