package com.zbx.kltopb;

import java.io.File;

/**
 * @日期 2022/10/7
 * @作者 zbx
 * @描述
 **/
public interface MatchReplace {

    void replace(File file, File outDir) throws Exception;
}
