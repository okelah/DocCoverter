package com.tencent.fm.convert.bean;

/**
 * Created by pengfeining on 2018/11/2 0002.
 */
public enum SourceFileType {
        DOC(0), XLS(1), PPT(2);
    
    private int value;
    
    SourceFileType(int value) {
        this.value = value;
    }
    
    public int getValue() {
        return value;
    }
    
    public static SourceFileType fromValue(int value) {
        for (SourceFileType sourceFileType : SourceFileType.values()) {
            if (sourceFileType.getValue() == value) {
                return sourceFileType;
            }
        }
        return SourceFileType.DOC;
    }
}
