package com.tencent.fm.convert;

import com.tencent.fm.convert.bean.SourceFile;
import com.tencent.fm.convert.bean.TargetFile;

/**
 * Created by pengfeining on 2018/11/2 0002.
 */
public interface FileConvert {
    public void convert(SourceFile sourceFile, TargetFile targetFile);
}
