package com.tencent.fm.convert.bean;

/**
 * Created by pengfeining on 2018/11/2 0002.
 */
public enum SourceFileType {
        DOC(0),DOCX(1), XLS(2),XLSX(3),PPT(4),PPTX(5);
    
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
