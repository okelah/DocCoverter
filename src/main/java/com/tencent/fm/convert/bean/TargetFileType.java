package com.tencent.fm.convert.bean;

/**
 * Created by pengfeining on 2018/11/2 0002.
 */
public enum TargetFileType {
        HTML(0), PDF(1);
    
    private int value;
    
    TargetFileType(int value) {
        this.value = value;
    }
    
    public int getValue() {
        return value;
    }
    
    public static TargetFileType fromValue(int value) {
        for (TargetFileType targetFileType : TargetFileType.values()) {
            if (targetFileType.getValue() == value) {
                return targetFileType;
            }
        }
        return TargetFileType.HTML;
    }
}
