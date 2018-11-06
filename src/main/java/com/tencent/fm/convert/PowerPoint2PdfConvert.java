package com.tencent.fm.convert;

import com.tencent.fm.convert.bean.SourceFile;
import com.tencent.fm.convert.bean.TargetFile;

/**
 * Created by pengfeining on 2018/11/2 0002.
 */
public interface PowerPoint2PdfConvert extends Convert {
    void powerpoint2pdf(SourceFile sourceFile, TargetFile targetFile);
    void ppt2pdf(SourceFile sourceFile,TargetFile targetFile);
    void pptx2pdf(SourceFile sourceFile,TargetFile targetFile);
}
