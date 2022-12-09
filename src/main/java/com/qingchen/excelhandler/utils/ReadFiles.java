package com.qingchen.excelhandler.utils;

import java.io.File;
import java.util.ArrayList;
import java.util.List;

/**
 * 读取指定文件夹下所有文件
 *
 * @author qingchen
 * @date 1/11/2022 下午 2:13
 */

public class ReadFiles {
    public static List<File> readFile(String fileDir) {
        List<File> fileList = new ArrayList<File>();
        File file = new File(fileDir);
        File[] files = file.listFiles();// 获取目录下的所有文件或文件夹
        if (files == null) {// 如果目录为空，直接退出
            return null;
        }

        // 遍历，目录下的所有文件
        for (File f : files) {
            // true当且仅当该抽象路径名表示的文件存在且为普通文件时
            if (f.isFile()) {
                fileList.add(f);
            } else if (f.isDirectory()) {
                System.out.println(f.getAbsolutePath());
                readFile(f.getAbsolutePath());
            }
        }
        return fileList;
    }

}
