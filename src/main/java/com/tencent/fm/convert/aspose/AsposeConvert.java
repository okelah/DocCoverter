package com.tencent.fm.convert.aspose;

import com.aspose.cells.*;
import com.aspose.words.*;
import com.aspose.words.HtmlSaveOptions;
import com.aspose.words.License;
import com.aspose.words.SaveFormat;
import com.tencent.fm.convert.Excel2HtmlConvert;
import com.tencent.fm.convert.Excel2PdfConvert;
import com.tencent.fm.convert.Word2HtmlConvert;
import com.tencent.fm.convert.Word2PdfConvert;
import com.tencent.fm.convert.bean.SourceFile;
import com.tencent.fm.convert.bean.TargetFile;
import org.apache.commons.lang.StringUtils;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Value;

import java.io.File;
import java.io.InputStream;


/**
 * Created by pengfeining on 2018/10/25 0025. 支持mhtml 如果文后缀名是mhtml
 */
public class AsposeConvert implements Word2HtmlConvert, Word2PdfConvert, Excel2HtmlConvert, Excel2PdfConvert {
    
    @Value("${xls.font.folder}") // xls常用字体包文件夹路径 没有用到
    String XlsFontFolder;
    
    @Value("${word.license.path}") // xls常用字体包文件夹路
    String wordLicensePath;
    
    @Value("${excel.license.path}") // xls常用字体包文件夹路
    String excelLicensePath;
    
    public void setXlsFontFolder(String xlsFontFolder) {
        XlsFontFolder = xlsFontFolder;
    }
    
    public void setLicensePath(String excelLicensePath) {
        this.excelLicensePath = excelLicensePath;
    }
    
    Logger logger = LoggerFactory.getLogger(AsposeConvert.class);
    
    /**
     * 获取license
     *
     * @return
     */
    // public static boolean getLicense() {
    // boolean result = false;
    // try {
    // InputStream is = AsposeConvert.class.getClassLoader().getResourceAsStream("aspose/Aspose.Total.Java.lic");
    // License aposeLic = new License();
    // aposeLic.setLicense(is);
    // result = true;
    // } catch (Exception e) {
    // e.printStackTrace();
    // }
    // return result;
    // }
    
    public boolean getWordLicense() {
        InputStream license;
        boolean result = false;
        try {
            license = AsposeConvert.class.getClassLoader().getResourceAsStream(wordLicensePath);
            License aposeLic = new License();
            aposeLic.setLicense(license);
            result = true;
        } catch (Exception e) {
            e.printStackTrace();
        }
        return result;
    }
    
    public boolean getExcelLicense() {
        InputStream license;
        boolean result = false;
        try {
            license = AsposeConvert.class.getClassLoader().getResourceAsStream(excelLicensePath);
            com.aspose.cells.License aposeLic = new com.aspose.cells.License();
            aposeLic.setLicense(license);
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
    public void word2pdf(SourceFile sourceFile, TargetFile targetFile) {
        switch (sourceFile.getSourceFileType()) {
        case DOC:
            doc2pdf(sourceFile, targetFile);
            break;
        case DOCX:
            docx2pdf(sourceFile, targetFile);
            break;
        default:
            logger.info("{} is not a world", sourceFile.getPath());
            break;
        }
    }
    
    @Override
    public void doc2pdf(SourceFile sourceFile, TargetFile targetFile) {
        try {
            String inputFilePath = sourceFile.getPath();
            String outputFilePath = targetFile.getPath();
            // 验证License
            if (!getWordLicense()) {
                return;
            }
            Document doc = new Document(inputFilePath);
            doc.save(outputFilePath, SaveFormat.PDF);// 全面支持DOC, DOCX, OOXML, RTF HTML, OpenDocument, PDF, EPUB, XPS, SWF 相互转换
            
        } catch (Exception e) {
            logger.error("asopose convert error:" + e.getMessage());
            e.printStackTrace();
        }
    }
    
    @Override
    public void docx2pdf(SourceFile sourceFile, TargetFile targetFile) {
        doc2pdf(sourceFile, targetFile);
    }
    
    @Override
    public void word2html(SourceFile sourceFile, TargetFile targetFile) {
        switch (sourceFile.getSourceFileType()) {
        case DOC:
            doc2html(sourceFile, targetFile);
            break;
        case DOCX:
            docx2html(sourceFile, targetFile);
            break;
        default:
            logger.info("{} is not a world", sourceFile.getPath());
            break;
        }
    }
    
    /**
     * 注意输出文件格式与输入文件格式 支持html 与mhtml的输出 mhtml 目前存在问题
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
            if (!getWordLicense()) {
                return;
            }
            
            File inputFile = new File(inputFilePath);
            if (!inputFile.exists() || !inputFile.isFile()) {
                logger.error("asopose params inputFilePath error:请设置为完整的文件路径");
                return;
            }
            
            boolean mhtmlFormat = false;
            if (outputFilePath.endsWith("mhtml")) {
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
    
    @Override
    public void docx2html(SourceFile sourceFile, TargetFile targetFile) {
        doc2html(sourceFile, targetFile);
    }
    
    @Override
    public void excel2html(SourceFile sourceFile, TargetFile targetFile) {
        switch (sourceFile.getSourceFileType()) {
        case XLS:
            xls2html(sourceFile, targetFile);
            break;
        case XLSX:
            xlsx2html(sourceFile, targetFile);
            break;
        default:
            logger.info("{} is not a excel", sourceFile.getPath());
            break;
        }
    }
    
    @Override
    public void xls2html(SourceFile sourceFile, TargetFile targetFile) {
        try {
            
            String inputFilePath = sourceFile.getPath();
            String outputFilePath = targetFile.getPath();
            // 验证License
            if (!getExcelLicense()) {
                return;
            }
            
            File inputFile = new File(inputFilePath);
            if (!inputFile.exists() || !inputFile.isFile()) {
                logger.error("asopose params inputFilePath error:请设置为完整的文件路径");
                return;
            }
            
            logger.info("Aspoese.Cells for Java Version:" + CellsHelper.getVersion());
            Workbook workbook = new Workbook(inputFilePath);
            com.aspose.cells.HtmlSaveOptions options = new com.aspose.cells.HtmlSaveOptions();
            // 隐藏压盖的内容
            // 破解版本低问题 注释掉18.10才有
            // options.setHtmlCrossStringType(HtmlCrossType.CROSS_HIDE_RIGHT);
            //
            // options.setExportDocumentProperties(false);
            // options.setExportWorkbookProperties(false);
            // options.setExportWorksheetProperties(false);
            //
            // options.setExportSimilarBorderStyle(true);
            //
            // options.setExportWorksheetCSSSeparately(true);
            //
            // options.setTableCssId("tableCssId");
            // //去掉没用的
            // options.setExcludeUnusedStyles(true);
            workbook.save(outputFilePath, options);
            
        } catch (Exception e) {
            logger.error("asopose convert error:" + e.getMessage());
            e.printStackTrace();
        }
    }
    
    @Override
    public void xlsx2html(SourceFile sourceFile, TargetFile targetFile) {
        xls2html(sourceFile, targetFile);
    }
    
    @Override
    public void excel2pdf(SourceFile sourceFile, TargetFile targetFile) {
        switch (sourceFile.getSourceFileType()) {
        case XLS:
            xls2pdf(sourceFile, targetFile);
            break;
        case XLSX:
            xlsx2pdf(sourceFile, targetFile);
            break;
        default:
            logger.info("{} is not a excel", sourceFile.getPath());
            break;
        }
    }
    
    /**
     * 注意所有引用的类需要为cells的
     * 
     * @param sourceFile
     * @param targetFile
     */
    @Override
    public void xls2pdf(SourceFile sourceFile, TargetFile targetFile) {
        try {
            String inputFilePath = sourceFile.getPath();
            String outputFilePath = targetFile.getPath();
            // 验证License
            if (!getExcelLicense()) {
                return;
            }
            
            File inputFile = new File(inputFilePath);
            if (!inputFile.exists() || !inputFile.isFile()) {
                logger.error("asopose params inputFilePath error:请设置为完整的文件路径");
                return;
            }
            Workbook workbook = null;
            
            if (!StringUtils.isEmpty(XlsFontFolder)) {
                // IndividualFontConfigs fontConfigs=new IndividualFontConfigs();
                // fontConfigs.setFontFolder("",false);
                // com.aspose.cells.LoadOptions loadOptions=new com.aspose.cells.LoadOptions(com.aspose.cells.LoadFormat.XLSX);
                //// loadOptions.setFontConfigs(fontConfigs);
                //
                // workbook =new Workbook(inputFilePath,loadOptions);
            } else {
                workbook = new Workbook(inputFilePath);
            }
            
            com.aspose.cells.PdfSaveOptions options = new com.aspose.cells.PdfSaveOptions(com.aspose.cells.SaveFormat.PDF);
            
            options.setAllColumnsInOnePagePerSheet(true);
            
            // options.setCustomPropertiesExport(PdfCustomPropertiesExport.STANDARD);
            
            workbook.save(outputFilePath, options);
            
        } catch (Exception e) {
            logger.error("asopose convert error:" + e.getMessage());
            e.printStackTrace();
        }
    }
    
    @Override
    public void xlsx2pdf(SourceFile sourceFile, TargetFile targetFile) {
        xls2pdf(sourceFile, targetFile);
    }
}
