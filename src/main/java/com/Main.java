package com;
import com.aspose.words.Document;
import com.aspose.words.HtmlMetafileFormat;
import com.aspose.words.HtmlSaveOptions;
import com.tencent.fm.convert.*;
import com.tencent.fm.convert.aspose.AsposeConvert;
import com.tencent.fm.convert.bean.SourceFile;
import com.tencent.fm.convert.bean.TargetFile;
import com.tencent.fm.convert.docx4j.Docx4jConvert;
import com.tencent.fm.convert.poi.PoiConvert;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;


/**
 * Created by pengfeining on 2018/10/23 0023.
 */
public class Main {

    static Logger logger = LoggerFactory.getLogger(Main.class);

    public static void main(String[] args) throws Exception {

//        String testfile="my.docx";
//        String testfile = "mac-edit.docx";
//        String testfile = "chart.docx";
//        doc4jHtmlTest(testfile);
//        doc4jHtmlTest(testfile);
//        asposeHtmlTest(testfile);
//        asposeHtmlTest(testfile);
//        saveHtmlWithMetafileFormat("E:\\ConvertTest\\");
//        doc4jPdfTest(testfile);
//        doc4jPdfTest(testfile);
//        asposePdfTest(testfile);
//        asposePdfTest(testfile);
//        pdftest();
//        htmlTest();
        xlsHtml();
//        xlsHtml();
//        xlsPdf();
    }


    public static void  pdftest(){
        SourceFile sourceFile=new SourceFile("//Users/null//Desktop//手游报告-Part2-1103v5.docx");
        TargetFile targetFile=new TargetFile("//Users/null//Desktop//doc4j.pdf");
//        Word2PdfConvert doc2PdfConvert=new AsposeConvert();
        Word2PdfConvert doc2PdfConvert=new Docx4jConvert();
        doc2PdfConvert.doc2pdf(sourceFile,targetFile);
    }

    public static void htmlTest() throws Exception{
        SourceFile sourceFile=new SourceFile("//Users//null//Desktop//手游报告-Part2-1103v5.docx");
        TargetFile targetFile=new TargetFile("//Users//null//Desktop//poi.html");
        Word2HtmlConvert word2HtmlConvert=new PoiConvert();
        word2HtmlConvert.word2html(sourceFile,targetFile);
    }

    public static void xlsHtml() throws Exception{
        SourceFile sourceFile=new SourceFile("//Users//null//Desktop//公司对比.xlsx");
        TargetFile targetFile=new TargetFile("//Users//null//Desktop//asposeXls.html");
//        Document document=new Document(sourceFile.getPath());
//        document.save(targetFile.getPath(),SaveFormat.MHTML);
//        Word2HtmlConvert doc2HtmlConvert=new AsposeConvert();
        Excel2HtmlConvert excel2HtmlConvert=new PoiConvert();
        excel2HtmlConvert.excel2html(sourceFile,targetFile);
    }


    public static void xlsPdf() throws Exception{
        SourceFile sourceFile=new SourceFile("//Users//null//Desktop//sampleSpecifyIndividualOrPrivateSetOfFontsForWorkbookRendering.xlsx");
        TargetFile targetFile=new TargetFile("//Users//null//Desktop//asposePdf.pdf");
//        Document document=new Document(sourceFile.getPath());
//        document.save(targetFile.getPath(),SaveFormat.MHTML);
//        Word2HtmlConvert doc2HtmlConvert=new AsposeConvert();
        Excel2PdfConvert xls2PdfConvert=new AsposeConvert();
        xls2PdfConvert.xls2pdf(sourceFile,targetFile);
    }


    public static void saveHtmlWithMetafileFormat(String dataDir) throws Exception
    {
        // ExStart:SaveHtmlWithMetafileFormat
        Document doc = new Document(dataDir + "mac_edit.docx");
        HtmlSaveOptions options = new HtmlSaveOptions();
        options.setMetafileFormat(HtmlMetafileFormat.EMF_OR_WMF);

        dataDir = dataDir + "SaveHtmlWithMetafileFormat_out.html";
        doc.save(dataDir, options);
        // ExEnd:SaveHtmlWithMetafileFormat
        System.out.println("\nDocument saved with Metafile format.\nFile saved at " + dataDir);
    }

    static void doc4jPdfTest(String filename) throws Exception {
        String pathInput = "E:\\ConvertTest\\" + filename;
        String pathOutput = "E:\\ConvertTest\\docx4j\\" + filename.substring(0, filename.indexOf(".")) + ".pdf";
//        String pathInput = "/Users/null/Documents/SelfStudy/" + filename;
//        String pathOutput = "/Users/null/Documents/SelfStudy/" + filename.substring(0, filename.indexOf(".")) + ".html";

        long doc4jStartTime = System.currentTimeMillis();
        // doc4jConvert
        FileInputStream fileInputStream = new FileInputStream(new File(pathInput));
        FileOutputStream fileOutputStream = new FileOutputStream(new File(pathOutput));
//        Convert doc4jConvert = new Doc4jConvert();
//        doc4jConvert.convertDoc2Pdf(fileInputStream, fileOutputStream);
//        doc4jConvert.convertDoc2Pdf(pathInput, pathOutput);
        long doc4jEndTime = System.currentTimeMillis();
        logger.info("docx4j cost {} ms", doc4jEndTime - doc4jStartTime);
    }

    static void asposePdfTest(String filename) throws Exception {
//        String pathInput = "E:\\ConvertTest\\" + filename;
//        String pathOutput = "E:\\ConvertTest\\aspose\\" + filename.substring(0, filename.indexOf(".")) + ".pdf";
        String pathInput = "/Users/null/Documents/SelfStudy/" + filename;
        String pathOutput = "/Users/null/Documents/SelfStudy/" + filename.substring(0, filename.indexOf(".")) + ".pdf";

        // asposeConvert
        long asposeStartTime = System.currentTimeMillis();
        FileInputStream fileInputStream = new FileInputStream(new File(pathInput));
        FileOutputStream fileOutputStream = new FileOutputStream(new File(pathOutput));

        Convert asposeConvert = new AsposeConvert();
//        asposeConvert.convertDoc2Pdf(fileInputStream, fileOutputStream);
        long asposeEndTime = System.currentTimeMillis();
        logger.info("aspose cost {} ms", asposeEndTime - asposeStartTime);
    }

    static void doc4jHtmlTest(String filename) throws Exception {
//        String pathInput = "E:\\ConvertTest\\" + filename;
//        String pathOutput = "E:\\ConvertTest\\docx4j\\" + filename.substring(0, filename.indexOf(".")) + ".html";
        String pathInput = "/Users/null/Documents/SelfStudy/" + filename;
        String pathOutput = "/Users/null/Documents/SelfStudy/" + filename.substring(0, filename.indexOf(".")) + ".html";

        long doc4jStartTime = System.currentTimeMillis();
        FileInputStream fileInputStream = new FileInputStream(new File(pathInput));
        FileOutputStream fileOutputStream = new FileOutputStream(new File(pathOutput));
//        Convert convert = new Doc4jConvert();
//        convert.convertDoc2Html(fileInputStream, fileOutputStream);
        long doc4jEndTime = System.currentTimeMillis();
        logger.info("aspose cost {} ms", doc4jStartTime - doc4jEndTime);
    }

    static void asposeHtmlTest(String filename) throws Exception {
//        String pathInput = "E:\\ConvertTest\\" + filename;
//        String pathOutput = "E:\\ConvertTest\\docx4j\\" + filename.substring(0, filename.indexOf(".")) + ".html";
        String pathInput = "/Users/null/Documents/SelfStudy/" + filename;
        String pathOutput = "/Users/null/Documents/SelfStudy/" + filename.substring(0, filename.indexOf(".")) + ".html";

        long doc4jStartTime = System.currentTimeMillis();
        FileInputStream fileInputStream = new FileInputStream(new File(pathInput));
        FileOutputStream fileOutputStream = new FileOutputStream(new File(pathOutput));
        Convert convert = new AsposeConvert();
//        convert.convertDoc2Html(fileInputStream, fileOutputStream);
        long doc4jEndTime = System.currentTimeMillis();
        logger.info("aspose cost {} ms", doc4jStartTime - doc4jEndTime);
    }
}
