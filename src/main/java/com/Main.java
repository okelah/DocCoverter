package com;

import com.tencent.fm.convert.bean.SourceFile;
import com.tencent.fm.convert.bean.TargetFile;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.context.ApplicationContext;
import org.springframework.context.support.ClassPathXmlApplicationContext;

import com.tencent.fm.convert.FileConvert;


/**
 * Created by pengfeining on 2018/10/23 0023.
 */
public class Main {
    
    static Logger logger = LoggerFactory.getLogger(Main.class);
    
    public static void main(String[] args) throws Exception {
        
        ApplicationContext ac = new ClassPathXmlApplicationContext("applicationContext.xml");
        FileConvert fileConvert = (FileConvert) ac.getBean("fileConvert");
//        fileConvert.convert(new SourceFile("E:\\test.docx"), new TargetFile("E:\\test.html"));
//        fileConvert.convert(new SourceFile("E:\\test.docx"), new TargetFile("E:\\test.pdf"));
        fileConvert.convert(new SourceFile("E:\\test.xlsx"), new TargetFile("E:\\test1.pdf"));
//        fileConvert.convert(new SourceFile("E:\\test.xlsx"), new TargetFile("E:\\test1.html"));
    }
}
