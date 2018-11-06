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
        switch (extensionName){
            case "html": targetFileType=TargetFileType.HTML;break;
            case "pdf":targetFileType=TargetFileType.PDF;break;
            default:break;
        }
        return targetFileType;
    }

    public void setTargetFileType(TargetFileType targetFileType) {
        this.targetFileType = targetFileType;
    }
}
