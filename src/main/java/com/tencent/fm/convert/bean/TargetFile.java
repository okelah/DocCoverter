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
        if(targetFileType!=null){
            return targetFileType;
        }
        String extensionName= getExtensionName();
        if ("html".equals(extensionName)) {
            targetFileType = TargetFileType.HTML;
        } else if ("pdf".equals(extensionName)) {
            targetFileType = TargetFileType.PDF;
        } else {
        }
        return targetFileType;
    }

    public void setTargetFileType(TargetFileType targetFileType) {
        this.targetFileType = targetFileType;
    }
}
