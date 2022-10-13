package com.zbx.kltopb.impl;

import cn.hutool.core.io.FileUtil;
import cn.hutool.core.io.IoUtil;
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
import java.text.DecimalFormat;
import java.util.*;

/**
 * @日期 2022/10/7
 * @作者 zbx
 * @描述
 **/
public class ExcelMatchReplace implements MatchReplace {

    private final DecimalFormat format = new DecimalFormat("#.#");

    @Override
    public void replace(File file, File outDir) throws Exception{
        File copy = FileUtil.copy(file, new File(outDir.getAbsolutePath() + File.separator + file.getName()), true);
        ExcelWriter writer = ExcelUtil.getWriter(copy);
        List<Sheet> sheets = writer.getSheets();
        try {
            for (Sheet sheet : sheets) {
                writer.setSheet(sheet.getSheetName());
                for (Row row : sheet) {
                    // 这两行颗粒转PB
                    replacePB(row.getCell(9), writer);
                    replacePB(row.getCell(27), writer);
                    // 根据封边计算开料长宽
                    calcDecrease(row, writer);
                    // 华广9mm以下板件开料长宽覆盖成品长宽
                    change9Line(row, writer);
                }
            }
            writer.flush();
            writer.close();
            boolean b = file.delete();
            if (!b) Console.log("原始文件 [{}] 删除失败!", file.getName());
        } catch (Exception e) {
            IoUtil.close(writer);
            copy.delete();
            throw e;
        }
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
        // 获取封边列
        Cell cell = row.getCell(19);
        if (ObjectUtil.isNull(cell)) return;
        // 封边信息
        String edge = cell.getStringCellValue();
        if (StrUtil.isNotEmpty(edge) && StrUtil.length(edge) == 4) {
            // 封边信息加星
            writer.getCell(cell.getColumnIndex(), cell.getRowIndex()).setCellValue(edge + " ★");
            // 厚边减尺
            if (edge.contains("2")) {
                double length = Double.parseDouble(row.getCell(14).getStringCellValue());
                double width = Double.parseDouble(row.getCell(15).getStringCellValue());
                length = length - (edge.charAt(0) == '2' ? 0.6 : 0) - (edge.charAt(1) == '2' ? 0.6 : 0);
                width = width - (edge.charAt(2) == '2' ? 0.6 : 0) - (edge.charAt(3) == '2' ? 0.6 : 0);
                writer.getCell(11, row.getRowNum()).setCellValue(format.format(length));
                writer.getCell(12, row.getRowNum()).setCellValue(format.format(width));
            }
        }
    }

    /**
     * 9厚度及以下覆盖开料长宽覆盖成品长宽
     * @param row 表格行
     * @param writer 写入流
     */
    public void change9Line(Row row, ExcelWriter writer) {
        // 获取厚度列
        Cell cell = row.getCell(16);
        if (ObjectUtil.isNull(cell)) return;
        // 厚度信息
        String thick = cell.getStringCellValue();
        if (StrUtil.isNotEmpty(thick) && thick.matches("^0?[1-9]$")) {
            String cutLength = row.getCell(11).getStringCellValue();
            String cutWidth = row.getCell(12).getStringCellValue();
            writer.getCell(14, row.getRowNum()).setCellValue(cutLength);
            writer.getCell(15, row.getRowNum()).setCellValue(cutWidth);
        }
    }

}
