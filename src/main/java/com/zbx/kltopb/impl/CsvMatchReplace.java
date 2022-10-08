package com.zbx.kltopb.impl;

import cn.hutool.core.lang.Console;
import cn.hutool.core.text.csv.*;
import cn.hutool.core.util.StrUtil;
import com.zbx.kltopb.MatchReplace;

import java.io.File;
import java.io.FileInputStream;
import java.io.InputStreamReader;
import java.nio.charset.Charset;

/**
 * @日期 2022/10/7
 * @作者 zbx
 * @描述
 **/
public class CsvMatchReplace implements MatchReplace {

    @Override
    public void replace(File file, File outDir) throws Exception {
        CsvReader reader = CsvUtil.getReader(new InputStreamReader(new FileInputStream(file), "gbk"));
        CsvWriter writer = CsvUtil.getWriter(outDir.getAbsolutePath() + File.separator + file.getName(), Charset.forName("gbk"));

        CsvData read = reader.read();
        for (CsvRow row : read.getRows()) {
            if (row.size() < 27) continue;
            replacePB(row, 9);
            replacePB(row, 27);
            calcDecrease(row);
        }
        writer.write(read);
        writer.flush();
        writer.close();
        reader.close();
        boolean b = file.delete();
        if (!b) Console.log("原始文件 [{}] 删除失败!", file.getName());
    }

    /**
     * 修改颗粒板为PB
     * @param row csv行
     * @param rowNum csv列号
     */
    public void replacePB(CsvRow row, int rowNum) {
        String value = row.get(rowNum);
        if (!value.contains("颗粒")) return;
        String str = value.replaceAll("E0实木颗粒", "PB").replaceAll("颗粒", "PB");
        row.set(rowNum, str);
    }

    /**
     * 计算开料长宽
     * @param row csv行
     */
    public void calcDecrease(CsvRow row) {
        String edge = row.get(19);
        if (StrUtil.length(edge) == 4 && edge.contains("2")) {
            double length = Double.parseDouble(row.get(14));
            double width = Double.parseDouble(row.get(15));
            length = length - (edge.charAt(0) == '2' ? 0.6 : 0) - (edge.charAt(1) == '2' ? 0.6 : 0);
            width = width - (edge.charAt(2) == '2' ? 0.6 : 0) - (edge.charAt(3) == '2' ? 0.6 : 0);
            row.set(11, String.valueOf(length));
            row.set(12, String.valueOf(width));
        }
    }
}

