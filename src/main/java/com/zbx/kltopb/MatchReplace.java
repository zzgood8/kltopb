package com.zbx.kltopb;

import java.io.File;
import java.util.ArrayList;
import java.util.List;

/**
 * @日期 2022/10/7
 * @作者 zbx
 * @描述
 **/
public interface MatchReplace {

    List<String> COLORS = new ArrayList<>();

    void replace(File file, File outDir) throws Exception;
}
