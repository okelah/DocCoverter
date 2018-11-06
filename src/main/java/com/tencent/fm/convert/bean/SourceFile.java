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
        if(sourceFileType!=null){
            return sourceFileType;
        }
       String extensionName= getExtensionName();
        switch (extensionName){
            case "doc":sourceFileType=SourceFileType.DOC;break;
            case "docx": sourceFileType=SourceFileType.DOCX;break;
            case "xls":sourceFileType=SourceFileType.XLS;break;
            case "xlsx":sourceFileType=SourceFileType.XLSX;break;
            case "ppt":sourceFileType=SourceFileType.PPT;break;
            case "pptx":sourceFileType=SourceFileType.PPTX;break;
            default:break;
        }
        return sourceFileType;
    }

    public void setSourceFileType(SourceFileType sourceFileType) {
        this.sourceFileType = sourceFileType;
    }
}
