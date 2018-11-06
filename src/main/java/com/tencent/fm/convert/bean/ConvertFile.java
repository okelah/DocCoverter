package com.tencent.fm.convert.bean;

import com.sun.istack.NotNull;
import org.apache.commons.lang.StringUtils;

import java.io.File;

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
        if(StringUtils.isEmpty(filename)&&!StringUtils.isEmpty(path)){
            File file=new File(path);
            filename=file.getName();
        }
        return filename;
    }

    public void setFilename(String filename) {
        this.filename = filename;
    }

    public String getDir() {
        if(StringUtils.isEmpty(dir)&&!StringUtils.isEmpty(path)){
            File file=new File(path);
            dir= path.substring(0,path.length()-file.getName().length());
        }
        return dir;
    }

    public void setDir(String dir) {
        this.dir = dir;
    }


    public  String getExtensionName() {
        //path 获取filename
        String filename=getFilename();
        if ((filename != null) && (filename.length() > 0)) {
            int dot = filename.lastIndexOf('.');
            if ((dot > -1) && (dot < (filename.length() - 1))) {
                return filename.substring(dot + 1);
            }
        }
        return filename;
    }
}
