package com.tencent.fm.convert.bean;

/**
 * Created by pengfeining on 2018/11/2 0002.
 */
public class TargetFile extends ConvertFile {
    private TargetFileType targetFileType;

    public TargetFile(String path) {
        super(path);
    }

    public TargetFileType getTargetFileType() {
        return targetFileType;
    }

    public void setTargetFileType(TargetFileType targetFileType) {
        this.targetFileType = targetFileType;
    }
}
