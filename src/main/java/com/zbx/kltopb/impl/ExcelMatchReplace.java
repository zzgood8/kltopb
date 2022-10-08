package com.zbx.kltopb.impl;

import cn.hutool.core.io.FileUtil;
import cn.hutool.core.lang.Console;
import cn.hutool.core.util.ObjectUtil;
import cn.hutool.core.util.StrUtil;
import cn.hutool.poi.excel.ExcelUtil;
import cn.hutool.poi.excel.ExcelWriter;
import com.zbx.kltopb.MatchReplace;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import java.io.File;
import java.util.*;

/**
 * @日期 2022/10/7
 * @作者 zbx
 * @描述
 **/
public class ExcelMatchReplace implements MatchReplace {


    @Override
    public void replace(File file, File outDir) throws Exception{
        File copy = FileUtil.copy(file, new File(outDir.getAbsolutePath() + File.separator + file.getName()), true);
        ExcelWriter writer = ExcelUtil.getWriter(copy);
        List<Sheet> sheets = writer.getSheets();
        for (Sheet sheet : sheets) {
            writer.setSheet(sheet.getSheetName());
            for (Row row : sheet) {
                // 这两行颗粒转PB
                replacePB(row.getCell(9), writer);
                replacePB(row.getCell(27), writer);
                // 根据封边计算开料长宽
                calcDecrease(row, writer);
            }
        }
        writer.flush();
        writer.close();
        boolean b = file.delete();
        if (!b) Console.log("原始文件 [{}] 删除失败!", file.getName());
    }

    /**
     * 颗粒替换成PB
     * @param cell 单元格
     * @param writer 写入流
     */
    public void replacePB(Cell cell, ExcelWriter writer) {
        if (ObjectUtil.isNull(cell)) return;
        String value = cell.getStringCellValue();
        if (!value.contains("颗粒")) return;
        String str = value.replaceAll("E0实木颗粒", "PB").replaceAll("颗粒", "PB");
        writer.getCell(cell.getColumnIndex(), cell.getRowIndex()).setCellValue(str);
    }

    /**
     * 计算开料长宽
     * @param row 表格行
     * @param writer 写入流
     */
    public void calcDecrease(Row row, ExcelWriter writer) {
        Cell cell = row.getCell(19);
        if (ObjectUtil.isNull(cell)) return;
        String edge = cell.getStringCellValue();
        if (StrUtil.isNotEmpty(edge)) {
            writer.getCell(cell.getColumnIndex(), cell.getRowIndex()).setCellValue(edge + " ★");
        }
        if (StrUtil.length(edge) == 4 && edge.contains("2")) {
            double length = Double.parseDouble(row.getCell(14).getStringCellValue());
            double width = Double.parseDouble(row.getCell(15).getStringCellValue());
            length = length - (edge.charAt(0) == '2' ? 0.6 : 0) - (edge.charAt(1) == '2' ? 0.6 : 0);
            width = width - (edge.charAt(2) == '2' ? 0.6 : 0) - (edge.charAt(3) == '2' ? 0.6 : 0);
            writer.getCell(11, row.getRowNum()).setCellValue(length);
            writer.getCell(12, row.getRowNum()).setCellValue(width);
        }
    }
}
