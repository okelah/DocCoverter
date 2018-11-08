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
        if ("doc".equals(extensionName)) {
            sourceFileType = SourceFileType.DOC;
        } else if ("docx".equals(extensionName)) {
            sourceFileType = SourceFileType.DOCX;
        } else if ("xls".equals(extensionName)) {
            sourceFileType = SourceFileType.XLS;
        } else if ("xlsx".equals(extensionName)) {
            sourceFileType = SourceFileType.XLSX;
        } else if ("ppt".equals(extensionName)) {
            sourceFileType = SourceFileType.PPT;
        } else if ("pptx".equals(extensionName)) {
            sourceFileType = SourceFileType.PPTX;
        } else {
        }
        return sourceFileType;
    }

    public void setSourceFileType(SourceFileType sourceFileType) {
        this.sourceFileType = sourceFileType;
    }
}
