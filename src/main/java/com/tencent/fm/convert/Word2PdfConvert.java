package com.tencent.fm.convert;

import com.tencent.fm.convert.bean.SourceFile;
import com.tencent.fm.convert.bean.TargetFile;
import org.apache.poi.ss.formula.functions.T;

/**
 * Created by pengfeining on 2018/11/2 0002.
 */
public interface Word2PdfConvert extends Convert {
    void word2pdf(SourceFile sourceFile, TargetFile targetFile);
    void doc2pdf(SourceFile sourceFile,TargetFile targetFile);
    void docx2pdf(SourceFile sourceFile, TargetFile targetFile);
}
