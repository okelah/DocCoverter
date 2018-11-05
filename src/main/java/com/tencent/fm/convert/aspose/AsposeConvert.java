package com.tencent.fm.convert.aspose;

import com.aspose.cells.CellsHelper;
import com.aspose.cells.PdfCustomPropertiesExport;
import com.aspose.cells.Workbook;
import com.aspose.words.*;
import com.aspose.cells.*;
import com.aspose.words.HtmlSaveOptions;
import com.aspose.words.License;
import com.aspose.words.LoadFormat;
import com.aspose.words.LoadOptions;
import com.aspose.words.PdfSaveOptions;
import com.aspose.words.SaveFormat;
import com.tencent.fm.convert.*;
import com.tencent.fm.convert.bean.SourceFile;
import com.tencent.fm.convert.bean.TargetFile;
import org.docx4j.wml.Tr;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.File;
import java.io.InputStream;


/**
 * Created by pengfeining on 2018/10/25 0025. 支持mhtml 如果文后缀名是mhtml
 */
public class AsposeConvert implements Doc2HtmlConvert, Doc2PdfConvert,Xls2HtmlConvert,Xls2PdfConvert{
    
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
            long startTime=System.currentTimeMillis();
            String inputFilePath = sourceFile.getPath();
            String outputFilePath = targetFile.getPath();
            // 验证License
            if (!getLicense()) {
                return;
            }
            Document doc = new Document(inputFilePath);
            doc.save(outputFilePath, SaveFormat.PDF);// 全面支持DOC, DOCX, OOXML, RTF HTML, OpenDocument, PDF, EPUB, XPS, SWF 相互转换
            long endTime=System.currentTimeMillis();
            logger.info(doc.getOriginalFileName()+" 转化成功,花费时间{}ms",endTime-startTime);
        } catch (Exception e) {
            logger.error("asopose convert error:" + e.getMessage());
            e.printStackTrace();
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
            long startTime=System.currentTimeMillis();
            String inputFilePath = sourceFile.getPath();
            String outputFilePath = targetFile.getPath();
            // 验证License
            if (!getLicense()) {
                return;
            }
            
            File inputFile = new File(inputFilePath);
            if (!inputFile.exists() ||!inputFile.isFile()) {
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
            long endTime=System.currentTimeMillis();
            logger.info(doc.getOriginalFileName()+" 转化成功,花费时间{}ms",endTime-startTime);
        } catch (Exception e) {
            logger.error("asopose convert error:" + e.getMessage());
            e.printStackTrace();
        }
    }

    @Override
    public void xls2html(SourceFile sourceFile, TargetFile targetFile) {
        try {
            long startTime=System.currentTimeMillis();
            String inputFilePath = sourceFile.getPath();
            String outputFilePath = targetFile.getPath();
            // 验证License
            if (!getLicense()) {
                return;
            }

            File inputFile = new File(inputFilePath);
            if (!inputFile.exists() ||!inputFile.isFile()) {
                logger.error("asopose params inputFilePath error:请设置为完整的文件路径");
                return;
            }

           logger.info("Aspoese.Cells for Java Version:"+CellsHelper.getVersion());
            Workbook workbook= new Workbook(inputFilePath);
            com.aspose.cells.HtmlSaveOptions options=new com.aspose.cells.HtmlSaveOptions();
            //隐藏压盖的内容
            options.setHtmlCrossStringType(HtmlCrossType.CROSS_HIDE_RIGHT);

            options.setExportDocumentProperties(false);
            options.setExportWorkbookProperties(false);
            options.setExportWorksheetProperties(false);

            options.setExportSimilarBorderStyle(true);

            options.setExportWorksheetCSSSeparately(true);

            options.setTableCssId("tableCssId");
            //去掉没用的
            options.setExcludeUnusedStyles(true);
            workbook.save(outputFilePath,options);
            long endTime=System.currentTimeMillis();
            logger.info(workbook.getFileName()+" 转化成功,花费时间{}ms",endTime-startTime);
        } catch (Exception e) {
            logger.error("asopose convert error:" + e.getMessage());
            e.printStackTrace();
        }
    }


    /**
     * 注意所有引用的类需要为cells的
     * @param sourceFile
     * @param targetFile
     */
    @Override
    public void xls2pdf(SourceFile sourceFile, TargetFile targetFile) {
        try {
            long startTime=System.currentTimeMillis();
            String inputFilePath = sourceFile.getPath();
            String outputFilePath = targetFile.getPath();
            // 验证License
            if (!getLicense()) {
                return;
            }

            File inputFile = new File(inputFilePath);
            if (!inputFile.exists() ||!inputFile.isFile()) {
                logger.error("asopose params inputFilePath error:请设置为完整的文件路径");
                return;
            }
            //TODO 导入字体包 将常用的字体打包
            IndividualFontConfigs fontConfigs=new IndividualFontConfigs();
            fontConfigs.setFontFolder("",false);
            com.aspose.cells.LoadOptions loadOptions=new com.aspose.cells.LoadOptions(com.aspose.cells.LoadFormat.XLSX);
            loadOptions.setFontConfigs(fontConfigs);

            Workbook workbook=new Workbook(inputFilePath,loadOptions);
            com.aspose.cells.PdfSaveOptions options=new  com.aspose.cells.PdfSaveOptions( com.aspose.cells.SaveFormat.PDF);

            options.setAllColumnsInOnePagePerSheet(true);

            options.setCustomPropertiesExport(PdfCustomPropertiesExport.STANDARD);

            workbook.save(outputFilePath,options);
            long endTime=System.currentTimeMillis();
            logger.info(workbook.getFileName()+" 转化成功,花费时间{}ms",endTime-startTime);
        }catch (Exception e){
            logger.error("asopose convert error:" + e.getMessage());
            e.printStackTrace();
        }
    }
}
