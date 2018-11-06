package com.tencent.fm.convert;

import com.tencent.fm.convert.bean.SourceFile;
import com.tencent.fm.convert.bean.TargetFile;

/**
 * Created by pengfeining on 2018/11/2 0002.
 */
public interface Word2HtmlConvert extends Convert {
    void word2html(SourceFile sourceFile, TargetFile targetFile);
    void doc2html(SourceFile sourceFile,TargetFile targetFile);
    void docx2html(SourceFile sourceFile,TargetFile targetFile);
}
