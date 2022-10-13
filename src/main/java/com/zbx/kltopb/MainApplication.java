package com.zbx.kltopb;

import cn.hutool.core.date.DateUtil;
import cn.hutool.core.date.TimeInterval;
import cn.hutool.core.io.FileUtil;
import cn.hutool.core.lang.Console;
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

    private final static MatchReplace csv = new CsvMatchReplace();
    private final static MatchReplace excel = new ExcelMatchReplace();

    public static void main(String[] args) {
        // 计时器
        TimeInterval timer = DateUtil.timer();

        if (ArrayUtil.isEmpty(args) || args.length != 2) {
            exit("程序参数错误!");
        }

        File inputDir = FileUtil.file(args[0]);
        File outputDir = FileUtil.file(args[1]);

        if (!inputDir.exists() || !inputDir.isDirectory()) exit("输入目录不存在或者不是文件夹");
        if (!outputDir.exists() || !outputDir.isDirectory()) exit("输出目录不存在或者不是文件夹");

        try {
            convert(inputDir, outputDir);
            Console.log("本次转换耗时: {} 秒", timer.interval() / 1000f);
        } catch (Exception e) {
            System.out.println("发生未处理异常!");
            e.printStackTrace();
            System.exit(-1);
        }
    }

    public static void convert(File in, File out) throws Exception {
        File[] files = in.listFiles();
        if (ArrayUtil.isEmpty(files)) return;
        for (File file : files) {
            // 文件夹递归调用
            if (file.isDirectory()) {
                File outDir = new File(out + File.separator + file.getName());
                if (!outDir.exists()) outDir.mkdir();
                convert(file, outDir);
                file.delete();
            }
            // 解析文件
            if (file.getName().toUpperCase().matches(".+\\.((XLS)|(XLSX)|(CSV))$")) {
                String fileType = file.getName().substring(file.getName().lastIndexOf(".") + 1);
                if (StrUtil.equalsIgnoreCase(fileType, "CSV")) {
                    csv.replace(file, out);
                } else {
                    excel.replace(file, out);
                }
            }
        }

    }

    public static void exit(String msg) {
        System.out.println(msg);
        System.exit(-1);
    }
}
