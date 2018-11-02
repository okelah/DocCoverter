package com.tencent.fm.convert.bean;

/**
 * Created by pengfeining on 2018/11/2 0002.
 */
public class ConvertFile {
    
    private String path;

    private String filename;
    
    private String dir;

    public ConvertFile(String path) {
        this.path = path;
    }

    public String getPath() {
        return path;
    }

    public void setPath(String path) {
        this.path = path;
    }

    public String getFilename() {
        return filename;
    }

    public void setFilename(String filename) {
        this.filename = filename;
    }

    public String getDir() {
        return dir;
    }

    public void setDir(String dir) {
        this.dir = dir;
    }
}
