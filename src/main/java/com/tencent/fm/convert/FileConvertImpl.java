package com.tencent.fm.convert;

import com.tencent.fm.convert.bean.SourceFile;
import com.tencent.fm.convert.bean.SourceFileType;
import com.tencent.fm.convert.bean.TargetFile;
import com.tencent.fm.convert.bean.TargetFileType;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;


/**
 * Created by pengfeining on 2018/11/2 0002.
 */
public class FileConvertImpl implements FileConvert {
    
    Logger logger = LoggerFactory.getLogger(FileConvertImpl.class);
    
    Word2PdfConvert word2PdfConvert;
    
    Word2HtmlConvert word2HtmlConvert;
    
    Excel2PdfConvert excel2PdfConvert;
    
    Excel2HtmlConvert excel2HtmlConvert;
    
    @Override
    public void convert(SourceFile sourceFile, TargetFile targetFile) {
        long startTime = System.currentTimeMillis();
        SourceFileType sourceFileType = sourceFile.getSourceFileType();
        TargetFileType targetFileType = targetFile.getTargetFileType();
        
        if ((sourceFileType == SourceFileType.DOC || sourceFileType == SourceFileType.DOCX)
            && targetFileType == TargetFileType.HTML) {// word2html
            word2HtmlConvert.word2html(sourceFile, targetFile);
        }
        
        if ((sourceFileType == SourceFileType.DOC || sourceFileType == SourceFileType.DOCX)
            && targetFileType == TargetFileType.PDF) {// word2pdf
            word2PdfConvert.word2pdf(sourceFile, targetFile);
        }
        
        if ((sourceFileType == SourceFileType.XLS || sourceFileType == SourceFileType.XLSX)
            && targetFileType == TargetFileType.HTML) {// excel2html
            excel2HtmlConvert.excel2html(sourceFile, targetFile);
        }
        
        if ((sourceFileType == SourceFileType.XLS || sourceFileType == SourceFileType.XLSX)
            && targetFileType == TargetFileType.PDF) {// excel2pdf
            excel2PdfConvert.excel2pdf(sourceFile, targetFile);
        }
        long endTime = System.currentTimeMillis();
        logger.info(sourceFile.getPath() + " --> " + targetFile.getPath() + " finish, cost {} ms", endTime - startTime);
    }
    
    public void setWord2PdfConvert(Word2PdfConvert word2PdfConvert) {
        this.word2PdfConvert = word2PdfConvert;
    }
    
    public void setWord2HtmlConvert(Word2HtmlConvert word2HtmlConvert) {
        this.word2HtmlConvert = word2HtmlConvert;
    }
    
    public void setExcel2PdfConvert(Excel2PdfConvert excel2PdfConvert) {
        this.excel2PdfConvert = excel2PdfConvert;
    }
    
    public void setExcel2HtmlConvert(Excel2HtmlConvert excel2HtmlConvert) {
        this.excel2HtmlConvert = excel2HtmlConvert;
    }
}
