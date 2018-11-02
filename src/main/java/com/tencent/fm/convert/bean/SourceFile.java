package com.tencent.fm.convert.bean;

/**
 * Created by pengfeining on 2018/11/2 0002.
 */
public class SourceFile extends ConvertFile {
    private SourceFileType sourceFileType;

    public SourceFile(String path) {
        super(path);
    }

    public SourceFileType getSourceFileType() {
        return sourceFileType;
    }

    public void setSourceFileType(SourceFileType sourceFileType) {
        this.sourceFileType = sourceFileType;
    }
}
