package com.tencent.fm.convert.poi;

import com.tencent.fm.convert.Doc2HtmlConvert;
import com.tencent.fm.convert.Ppt2HtmlConvert;
import com.tencent.fm.convert.Xls2HtmlConvert;
import com.tencent.fm.convert.bean.SourceFile;
import com.tencent.fm.convert.bean.TargetFile;
import org.apache.commons.io.FileUtils;
import org.apache.poi.hssf.converter.ExcelToHtmlConverter;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.converter.PicturesManager;
import org.apache.poi.hwpf.converter.WordToHtmlConverter;
import org.apache.poi.hwpf.usermodel.Picture;
import org.apache.poi.hwpf.usermodel.PictureType;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.w3c.dom.Document;

import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.transform.OutputKeys;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;
import java.io.*;
import java.util.List;


/**
 * Created by pengfeining on 2018/11/2 0002.
 */
public class PoiConvert implements Doc2HtmlConvert, Xls2HtmlConvert, Ppt2HtmlConvert {
    
    Logger logger = LoggerFactory.getLogger(PoiConvert.class);
    
    /**
     * 本质上是转图片
     * 换个远吗
     * 
     * @param sourceFile
     * @param targetFile
     */
    @Override
    public void ppt2html(SourceFile sourceFile, TargetFile targetFile) {

    }
    
    @Override
    public void xls2html(SourceFile sourceFile, TargetFile targetFile) {
        try {
            String inputFilePath = sourceFile.getPath();
            String outputFilePath = targetFile.getPath();
            String outputDirPath = targetFile.getDir();
            InputStream input = new FileInputStream(inputFilePath);
            HSSFWorkbook excelBook = new HSSFWorkbook(input);
            ExcelToHtmlConverter excelToHtmlConverter = new ExcelToHtmlConverter(
                DocumentBuilderFactory.newInstance().newDocumentBuilder().newDocument());
            excelToHtmlConverter.processWorkbook(excelBook);
            List pics = excelBook.getAllPictures();
            if (pics != null) {
                for (int i = 0; i < pics.size(); i++) {
                    Picture pic = (Picture) pics.get(i);
                    try {
                        pic.writeImageContent(new FileOutputStream(outputDirPath + pic.suggestFullFileName()));
                    } catch (FileNotFoundException e) {
                        e.printStackTrace();
                    }
                }
            }
            Document htmlDocument = excelToHtmlConverter.getDocument();
            ByteArrayOutputStream outStream = new ByteArrayOutputStream();
            DOMSource domSource = new DOMSource(htmlDocument);
            StreamResult streamResult = new StreamResult(outStream);
            TransformerFactory tf = TransformerFactory.newInstance();
            Transformer serializer = tf.newTransformer();
            serializer.setOutputProperty(OutputKeys.ENCODING, "utf-8");
            serializer.setOutputProperty(OutputKeys.INDENT, "yes");
            serializer.setOutputProperty(OutputKeys.METHOD, "html");
            serializer.transform(domSource, streamResult);
            outStream.close();
            
            String content = new String(outStream.toByteArray());
            
            FileUtils.writeStringToFile(new File(outputDirPath), content, "utf-8");
        } catch (Exception e) {
            logger.error("poi convert xls2html failed:" + e.getMessage());
            e.printStackTrace();
        }
    }
    
    @Override
    public void doc2html(SourceFile sourceFile, TargetFile targetFile) {
        try {
            String inputFilePath = sourceFile.getPath();
            String outputFilePath = targetFile.getPath();
            String outputDirPath = targetFile.getDir();
            InputStream input = new FileInputStream(inputFilePath);
            HWPFDocument wordDocument = new HWPFDocument(input);
            WordToHtmlConverter wordToHtmlConverter = new WordToHtmlConverter(
                DocumentBuilderFactory.newInstance().newDocumentBuilder().newDocument());
            wordToHtmlConverter.setPicturesManager(new PicturesManager() {
                
                public String savePicture(byte[] content, PictureType pictureType, String suggestedName, float widthInches,
                    float heightInches) {
                    return suggestedName;
                }
            });
            wordToHtmlConverter.processDocument(wordDocument);
            List pics = wordDocument.getPicturesTable().getAllPictures();
            if (pics != null) {
                for (int i = 0; i < pics.size(); i++) {
                    Picture pic = (Picture) pics.get(i);
                    try {
                        pic.writeImageContent(new FileOutputStream(outputDirPath + pic.suggestFullFileName()));
                    } catch (FileNotFoundException e) {
                        e.printStackTrace();
                    }
                }
            }
            Document htmlDocument = wordToHtmlConverter.getDocument();
            ByteArrayOutputStream outStream = new ByteArrayOutputStream();
            DOMSource domSource = new DOMSource(htmlDocument);
            StreamResult streamResult = new StreamResult(outStream);
            TransformerFactory tf = TransformerFactory.newInstance();
            Transformer serializer = tf.newTransformer();
            serializer.setOutputProperty(OutputKeys.ENCODING, "utf-8");
            serializer.setOutputProperty(OutputKeys.INDENT, "yes");
            serializer.setOutputProperty(OutputKeys.METHOD, "html");
            serializer.transform(domSource, streamResult);
            outStream.close();
            String content = new String(outStream.toByteArray());
            FileUtils.writeStringToFile(new File(outputFilePath), content, "utf-8");
        } catch (Exception e) {
            logger.error("poi convert doc2html failed:" + e.getMessage());
            e.printStackTrace();
        }
    }
    
    /**
     * 检查文件是否是PPT
     * 
     * @param file
     * @return
     */
    public boolean checkFile(File file) {
        boolean isppt = false;
        String filename = file.getName();
        String suffixname = null;
        if (filename != null && filename.indexOf(".") != -1) {
            suffixname = filename.substring(filename.indexOf("."));
            if (suffixname.equals(".ppt")) {
                isppt = true;
            }
            return isppt;
        } else {
            return isppt;
        }
    }
}
