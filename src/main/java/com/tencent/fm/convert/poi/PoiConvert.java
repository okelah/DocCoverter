package com.tencent.fm.convert.poi;

import com.tencent.fm.convert.Word2HtmlConvert;
import com.tencent.fm.convert.PowerPoint2HtmlConvert;
import com.tencent.fm.convert.Excel2HtmlConvert;
import com.tencent.fm.convert.Word2PdfConvert;
import com.tencent.fm.convert.bean.SourceFile;
import com.tencent.fm.convert.bean.SourceFileType;
import com.tencent.fm.convert.bean.TargetFile;
import org.apache.commons.io.FileUtils;
import org.apache.poi.hssf.converter.ExcelToHtmlConverter;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.converter.PicturesManager;
import org.apache.poi.hwpf.converter.WordToHtmlConverter;
import org.apache.poi.hwpf.usermodel.Picture;
import org.apache.poi.hwpf.usermodel.PictureType;
import org.apache.poi.xwpf.converter.core.BasicURIResolver;
import org.apache.poi.xwpf.converter.core.FileImageExtractor;
import org.apache.poi.xwpf.converter.core.FileURIResolver;
import org.apache.poi.xwpf.converter.pdf.PdfConverter;
import org.apache.poi.xwpf.converter.pdf.PdfOptions;
import org.apache.poi.xwpf.converter.xhtml.XHTMLConverter;
import org.apache.poi.xwpf.converter.xhtml.XHTMLOptions;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
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
public class PoiConvert implements Word2HtmlConvert, Excel2HtmlConvert ,Word2PdfConvert {
    
    Logger logger = LoggerFactory.getLogger(PoiConvert.class);


    @Override
    public void excel2html(SourceFile sourceFile, TargetFile targetFile) {
        switch (sourceFile.getSourceFileType()){
            case XLS:xls2html(sourceFile,targetFile);break;
            case XLSX:xlsx2html(sourceFile,targetFile);break;
            default:logger.info("{} is not a excel",sourceFile.getPath());break;
        }
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
    public void xlsx2html(SourceFile sourceFile, TargetFile targetFile) {
        String inputFilePath = sourceFile.getPath();
        String outputFilePath = targetFile.getPath();
        try {
            FileInputStream inputStream=new FileInputStream(new File(inputFilePath));
            FileOutputStream outputStream=new FileOutputStream(new File(outputFilePath));
            String excelHtml = ExcelToHtmlUtil.readExcelToHtml(inputStream, true);
            outputStream.write(excelHtml.getBytes("UTF8"));
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    @Override
    public void word2html(SourceFile sourceFile, TargetFile targetFile) {
        switch (sourceFile.getSourceFileType()){
            case DOC:doc2html(sourceFile,targetFile);break;
            case DOCX:docx2html(sourceFile,targetFile);break;
            default:logger.info("{} is not a word",sourceFile.getPath());break;
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
     * POI中并没有提供将DOCX文件转化为HTML、XML等格式的接口，这里采用的是XDocReport中提供的接口，此外XDocReport中还提供将DOCX直接转化为PDF的接口
     * @param sourceFile
     * @param targetFile
     */
    public void docx2html(SourceFile sourceFile,TargetFile targetFile){
        try {
            String inputFilePath = sourceFile.getPath();
            String outputFilePath = targetFile.getPath();
            String outputDirPath = targetFile.getDir();
            long startTime = System.currentTimeMillis();
            XWPFDocument document = new XWPFDocument(new FileInputStream(inputFilePath));
            XHTMLOptions options = XHTMLOptions.create().indent(4);
            // 导出图片
            File imageFolder = new File(outputDirPath);
            options.setExtractor(new FileImageExtractor(imageFolder));
            // URI resolver
            options.URIResolver(new FileURIResolver(imageFolder));
            File outFile = new File(outputFilePath);
            outFile.getParentFile().mkdirs();
            OutputStream out = new FileOutputStream(outFile);
            XHTMLConverter.getInstance().convert(document, out, options);
        }catch (Exception e){
            logger.error("poi convert docx2html failed:"+e.getMessage());
            e.printStackTrace();
        }
    }

    @Override
    public void word2pdf(SourceFile sourceFile, TargetFile targetFile) {
        switch (sourceFile.getSourceFileType()){
            case DOC:doc2pdf(sourceFile,targetFile);break;
            case DOCX:docx2pdf(sourceFile,targetFile);break;
            default:logger.info("{} is not a word",sourceFile.getPath());break;
        }
    }

    @Override
    public void doc2pdf(SourceFile sourceFile, TargetFile targetFile) {
        docx2pdf(sourceFile,targetFile);
    }

    @Override
    public void docx2pdf(SourceFile sourceFile, TargetFile targetFile) {
        try {
            String inputFilePath = sourceFile.getPath();
            String outputFilePath = targetFile.getPath();
            long startTime = System.currentTimeMillis();
        XWPFDocument document = new XWPFDocument(new FileInputStream(new File(inputFilePath)));

        // 2) Prepare Pdf options
        PdfOptions options = PdfOptions.create();

        // 3) Convert XWPFDocument to Pdf
        OutputStream out = new FileOutputStream(new File(outputFilePath));
        PdfConverter.getInstance().convert(document, out, options);
            long endTime = System.currentTimeMillis();}catch (Exception e){
            logger.error("poi convert docx2pdf failed:"+e.getMessage());
            e.printStackTrace();
        }
    }
}
