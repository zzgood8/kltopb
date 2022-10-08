package com.zbx.kltopb;

import cn.hutool.core.io.FileUtil;
import cn.hutool.core.util.ArrayUtil;
import cn.hutool.core.util.StrUtil;
import com.zbx.kltopb.impl.CsvMatchReplace;
import com.zbx.kltopb.impl.ExcelMatchReplace;

import java.io.File;

/**
 * @日期 2022/10/7
 * @作者 zbx
 * @描述
 **/
public class MainApplication {
    public static void main(String[] args) {

        if (ArrayUtil.isEmpty(args) || args.length != 2) {
            exit("程序参数错误!");
        }

        File inputDir = FileUtil.file(args[0]);
        File outputDir = FileUtil.file(args[1]);

        if (!inputDir.exists() || !inputDir.isDirectory()) exit("输入目录不存在或者不是文件夹");
        if (!outputDir.exists() || !outputDir.isDirectory()) exit("输出目录不存在或者不是文件夹");

        File[] files = inputDir.listFiles(pathname -> pathname.getName().toUpperCase().matches(".+\\.((XLS)|(XLSX)|(CSV))$"));
        if (ArrayUtil.isEmpty(files)) exit("输入文件夹为空");

        MatchReplace csv = new CsvMatchReplace();
        MatchReplace excel = new ExcelMatchReplace();

        assert files != null;
        for (File file : files) {
            String fileType = file.getName().substring(file.getName().lastIndexOf(".") + 1);
            try {
                if (StrUtil.equalsIgnoreCase(fileType, "CSV")) {
                    csv.replace(file, outputDir);
                } else {
                    excel.replace(file, outputDir);
                }
            }catch (Exception e) {
                System.out.println("发生未处理异常!");
                e.printStackTrace();
                System.exit(-1);
            }

        }
    }

    public static void exit(String msg) {
        System.out.println(msg);
        System.exit(-1);
    }
}
