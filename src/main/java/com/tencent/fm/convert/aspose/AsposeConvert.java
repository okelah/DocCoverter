package com.tencent.fm.convert.aspose;

import com.aspose.words.*;
import com.tencent.fm.convert.*;
import com.tencent.fm.convert.bean.SourceFile;
import com.tencent.fm.convert.bean.TargetFile;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.File;
import java.io.InputStream;


/**
 * Created by pengfeining on 2018/10/25 0025. 支持mhtml 如果文后缀名是mhtml
 */
public class AsposeConvert implements Doc2HtmlConvert, Doc2PdfConvert {
    
    Logger logger = LoggerFactory.getLogger(AsposeConvert.class);
    
    /**
     * 获取license
     *
     * @return
     */
    public static boolean getLicense() {
        boolean result = false;
        try {
            InputStream is = AsposeConvert.class.getClassLoader().getResourceAsStream("aspose/Aspose.Total.Java.lic");
            License aposeLic = new License();
            aposeLic.setLicense(is);
            result = true;
        } catch (Exception e) {
            e.printStackTrace();
        }
        return result;
    }
    
    /**
     * 判断是否为空
     *
     * @param obj
     *            字符串对象
     * @return
     */
    protected static boolean notNull(String obj) {
        if (obj != null && !obj.equals("") && !obj.equals("undefined") && !obj.trim().equals("") && obj.trim().length() > 0) {
            return true;
        }
        return false;
    }
    
    @Override
    public void doc2pdf(SourceFile sourceFile, TargetFile targetFile) {
        try {
            String inputFilePath = sourceFile.getPath();
            String outputFilePath = targetFile.getPath();
            // 验证License
            if (!getLicense()) {
                return;
            }
            Document doc = new Document(inputFilePath);
            doc.save(outputFilePath, SaveFormat.PDF);// 全面支持DOC, DOCX, OOXML, RTF HTML, OpenDocument, PDF, EPUB, XPS, SWF 相互转换
        } catch (Exception e) {
            logger.error("asopose convert error:" + e.getMessage());
            e.printStackTrace();
        }
    }
    
    /**
     * 注意输出文件格式与输入文件格式 支持html与xhtml的输出
     * 
     * @param sourceFile
     * @param targetFile
     */
    @Override
    public void doc2html(SourceFile sourceFile, TargetFile targetFile) {
        try {
            String inputFilePath = sourceFile.getPath();
            String outputFilePath = targetFile.getPath();
            // 验证License
            if (!getLicense()) {
                return;
            }
            
            File inputFile = new File(inputFilePath);
            if (inputFile.exists() && inputFile.isFile()) {
                logger.error("asopose params error:请设置为完整的文件路径");
                return;
            }
            // 默认名字加时间戳
            String fileName = inputFile.getName();
            
            boolean mhtmlFormat = false;
            if (fileName.endsWith("mhtml")) {
                mhtmlFormat = true;
            }
            
            Document doc = new Document(inputFilePath);
            
            if (mhtmlFormat) {// mhtml 将所有的资源打包成一个mhtm
                HtmlSaveOptions options = new HtmlSaveOptions();
                options.setSaveFormat(SaveFormat.MHTML);// 全面支持DOC, DOCX, OOXML, RTF HTML, OpenDocument, PDF, EPUB, XPS, SWF 相互转换
                options.setPrettyFormat(true);
                options.setExportCidUrlsForMhtmlResources(true);
                doc.save(outputFilePath, options);
            } else {// 在html的同级目录下 图片与css与字体做为一个单独的文件生成
                HtmlFixedSaveOptions options = new HtmlFixedSaveOptions();
                options.setSaveFontFaceCssSeparately(false);
                options.setPrettyFormat(true);
                doc.save(outputFilePath, options);
            }
        } catch (Exception e) {
            logger.error("asopose convert error:" + e.getMessage());
            e.printStackTrace();
        }
    }
}
