package com.tencent.fm.convert;

import com.tencent.fm.convert.bean.SourceFile;
import com.tencent.fm.convert.bean.TargetFile;

/**
 * Created by pengfeining on 2018/11/2 0002.
 */
public interface Excel2HtmlConvert extends Convert{
    void excel2html(SourceFile sourceFile, TargetFile targetFile);
    void xls2html(SourceFile sourceFile,TargetFile targetFile);
    void xlsx2html(SourceFile sourceFile,TargetFile targetFile);
}
