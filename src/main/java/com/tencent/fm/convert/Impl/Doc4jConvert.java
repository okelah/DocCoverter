package com.tencent.fm.convert.Impl;

import com.tencent.fm.convert.Convert;
import org.apache.commons.io.IOUtils;
import org.docx4j.Docx4J;
import org.docx4j.Docx4jProperties;
import org.docx4j.convert.out.ConversionFeatures;
import org.docx4j.convert.out.FOSettings;
import org.docx4j.convert.out.HTMLSettings;
import org.docx4j.convert.out.html.SdtToListSdtTagHandler;
import org.docx4j.convert.out.html.SdtWriter;
import org.docx4j.fonts.IdentityPlusMapper;
import org.docx4j.fonts.Mapper;
import org.docx4j.fonts.PhysicalFont;
import org.docx4j.fonts.PhysicalFonts;
import org.docx4j.model.fields.FieldUpdater;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.services.client.ConversionException;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.File;
import java.io.FileOutputStream;


/**
 * Created by pengfeining on 2018/10/23 0023.
 */
public class Doc4jConvert implements Convert {
    
    Logger logger = LoggerFactory.getLogger(Doc4jConvert.class);
    
    void checkFonts(String... fonts) {
        try {
            PhysicalFonts.discoverPhysicalFonts();
            for (String font:fonts) {
                PhysicalFont physicalFont = PhysicalFonts.get(font);
                if(physicalFont==null){
                    logger.warn("current system has not {} font",font);
                }
            }


        } catch (Exception e) {

        }
    }

    @Override
    public void convertDoc2Html(String inputFilePath, String outputFilePath) {
        boolean nestLists = true;
        
        try {
            WordprocessingMLPackage wordMLPackage = Docx4J.load(new File(inputFilePath));
            // HTML exporter setup (required)
            // .. the HTMLSettings object
            HTMLSettings htmlSettings = Docx4J.createHTMLSettings();
            
            htmlSettings.setImageDirPath("_files");
            htmlSettings.setImageTargetUri("_files");
            htmlSettings.setWmlPackage(wordMLPackage);
            
            /*
             * CSS reset, see http://itumbcom.blogspot.com.au/2013/06/css-reset-how-complex-it-should-be.html
             *
             * motivated by vertical space in tables in Firefox and Google Chrome.
             * 
             * If you have unwanted vertical space, in Chrome this may be coming from -webkit-margin-before and -webkit-margin-after
             * (in Firefox, margin-top is set to 1em in html.css)
             * 
             * Setting margin: 0 on p is enough to fix it.
             * 
             * See further http://www.css-101.org/articles/base-styles-sheet-for-webkit-based-browsers/
             */
            String userCSS = null;
            if (nestLists) {
                // use browser defaults for ol, ul, li
                userCSS = "html, body, div, span, h1, h2, h3, h4, h5, h6, p, a, img,  table, caption, tbody, tfoot, thead, tr, th, td "
                    + "{ margin: 0; padding: 0; border: 0;}" + "body {line-height: 1;} ";
            } else {
                userCSS = "html, body, div, span, h1, h2, h3, h4, h5, h6, p, a, img,  ol, ul, li, table, caption, tbody, tfoot, thead, tr, th, td "
                    + "{ margin: 0; padding: 0; border: 0;}" + "body {line-height: 1;} ";
                
            }
            htmlSettings.setUserCSS(userCSS);
            
            // Other settings (optional)
            // htmlSettings.setUserBodyTop("<H1>TOP!</H1>");
            // htmlSettings.setUserBodyTail("<H1>TAIL!</H1>");
            
            // Sample sdt tag handler (tag handlers insert specific
            // html depending on the contents of an sdt's tag).
            // This will only have an effect if the sdt tag contains
            // the string @class=XXX
            // SdtWriter.registerTagHandler("@class", new TagClass() );
            
            // SdtWriter.registerTagHandler(Containerization.TAG_BORDERS, new TagSingleBox() );
            // SdtWriter.registerTagHandler(Containerization.TAG_SHADING, new TagSingleBox() );
            
            // list numbering: depending on whether you want list numbering hardcoded, or done using <li>.
            if (nestLists) {
                SdtWriter.registerTagHandler("HTML_ELEMENT", new SdtToListSdtTagHandler());
            } else {
                htmlSettings.getFeatures().remove(ConversionFeatures.PP_HTML_COLLECT_LISTS);
            }
            
            // output to an OutputStream.
            
            // If you want XHTML output
            Docx4jProperties.setProperty("docx4j.Convert.Out.HTML.OutputMethodXML", true);
            
            FileOutputStream outputStream = new FileOutputStream(new File(outputFilePath));
            // Don't care what type of exporter you use
            // Docx4J.toHTML(htmlSettings, os, Docx4J.FLAG_NONE);
            // Prefer the exporter, that uses a xsl transformation
            Docx4J.toHTML(htmlSettings, outputStream, Docx4J.FLAG_EXPORT_PREFER_XSL);
            // Prefer the exporter, that doesn't use a xsl transformation (= uses a visitor)
            // Docx4J.toHTML(htmlSettings, os, Docx4J.FLAG_EXPORT_PREFER_NONXSL);
            
            // Clean up, so any ObfuscatedFontPart temp files can be deleted
            if (wordMLPackage.getMainDocumentPart().getFontTablePart() != null) {
                wordMLPackage.getMainDocumentPart().getFontTablePart().deleteEmbeddedFontTempFiles();
            }
            // This would also do it, via finalize() methods
            htmlSettings = null;
            wordMLPackage = null;
        } catch (Docx4JException e) {
            e.printStackTrace();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
    
    @Override
    public void convertDoc2Pdf(String inputFilePath, String outputFilePath) {
        /*
         * URL of converter instance; 如果有部署商业转换服务plutext可以直接采用 Docx4jProperties.setProperty("com.plutext.converter.URL", // ..
         * install your own at "http://localhost:9016/v1/00000000-0000-0000-0000-000000000000/convert"); // .. or perform a quick
         * test against //"https://converter-eval.plutext.com:443/v1/00000000-0000-0000-0000-000000000000/convert");
         */
        if (Docx4J.pdfViaFO()) {
            // Font regex (optional)
            // Set regex if you want to restrict to some defined subset of fonts
            // Here we have to do this before calling createContent,
            // since that discovers fonts
            String regex = null;
            // Windows:
            // String
            // regex=".*(calibri|camb|cour|arial|symb|times|Times|zapf).*";
            // regex=".*(calibri|camb|cour|arial|times|comic|georgia|impact|LSANS|pala|tahoma|trebuc|verdana|symbol|webdings|wingding).*";
            // Mac
            // String
            // regex=".*(Courier New|Arial|Times New Roman|Comic Sans|Georgia|Impact|Lucida Console|Lucida Sans Unicode|Palatino
            // Linotype|Tahoma|Trebuchet|Verdana|Symbol|Webdings|Wingdings|MS Sans Serif|MS Serif).*";
            PhysicalFonts.setRegex(regex);
        }
        // Document loading (required)
        WordprocessingMLPackage wordMLPackage;
        try {
            
            FileOutputStream outputStream = new FileOutputStream(new File(outputFilePath));
            // Load .docx or Flat OPC .xml
            wordMLPackage = WordprocessingMLPackage.load(new File(inputFilePath));
            
            // Refresh the values of DOCPROPERTY fields
            FieldUpdater updater = new FieldUpdater(wordMLPackage);
            updater.update(true);
            
            if (!Docx4J.pdfViaFO()) {
                // Since 3.3.0, Plutext's PDF Converter is used by default
                logger.info("Using Plutext's PDF Converter; add docx4j-export-fo if you don't want that");
                try {
                    Docx4J.toPDF(wordMLPackage, outputStream);
                } catch (Docx4JException e) {
                    if (e.getCause() != null) {
                        if (e.getCause() instanceof ConversionException) {
                            ConversionException ce = (ConversionException) e.getCause();
                            ce.printStackTrace();
                        }
                    } else {
                        IOUtils.closeQuietly(outputStream);
                        e.printStackTrace();
                    }
                    return;
                }
                return;
            }
            
            logger.info("Attempting to use XSL FO");
            // Set up font mapper (optional)

            Mapper fontMapper = new IdentityPlusMapper();
            fontMapper.put("等线", PhysicalFonts.get("DengXian Regular")); //字体别名的问题
            fontMapper.put("隶书", PhysicalFonts.get("LiSu"));
            fontMapper.put("宋体", PhysicalFonts.get("SimSun"));
            fontMapper.put("微软雅黑", PhysicalFonts.get("Microsoft Yahei"));
            fontMapper.put("黑体", PhysicalFonts.get("SimHei"));
            fontMapper.put("楷体", PhysicalFonts.get("KaiTi"));
            fontMapper.put("新宋体", PhysicalFonts.get("NSimSun"));
            fontMapper.put("华文行楷", PhysicalFonts.get("STXingkai"));
            fontMapper.put("华文仿宋", PhysicalFonts.get("STFangsong"));
            fontMapper.put("宋体扩展", PhysicalFonts.get("simsun-extB"));
            fontMapper.put("仿宋", PhysicalFonts.get("FangSong"));
            fontMapper.put("仿宋_GB2312", PhysicalFonts.get("FangSong_GB2312"));
            fontMapper.put("幼圆", PhysicalFonts.get("YouYuan"));
            fontMapper.put("华文宋体", PhysicalFonts.get("STSong"));
            fontMapper.put("华文中宋", PhysicalFonts.get("STZhongsong"));
            wordMLPackage.setFontMapper(fontMapper);
            
            FOSettings foSettings = Docx4J.createFOSettings();
            /*
             * 调试下需要
             */
            if (true) {
                foSettings.setFoDumpFile(new java.io.File("test.fo"));
            }
            foSettings.setWmlPackage(wordMLPackage);
            Docx4J.toFO(foSettings, outputStream, Docx4J.FLAG_EXPORT_PREFER_XSL);
            
            // 清空临时文件
            if (wordMLPackage.getMainDocumentPart().getFontTablePart() != null) {
                wordMLPackage.getMainDocumentPart().getFontTablePart().deleteEmbeddedFontTempFiles();
            }
            // This would also do it, via finalize() methods
            updater = null;
            foSettings = null;
            wordMLPackage = null;
            
        } catch (Docx4JException e) {
            e.printStackTrace();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    @Override
    public void convertXls2pdf(String inputFilePath, String outputFilePath) {

    }

    @Override
    public void convertXls2html(String inputFilePath, String outputFilePath) {

    }
}
