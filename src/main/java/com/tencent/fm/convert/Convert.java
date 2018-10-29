package com.tencent.fm.convert;

/**
 * Created by pengfeining on 2018/10/23 0023.
 */
public interface Convert {
    
    void convertDoc2Html(String inputFilePath, String outputFilePath);
    
    void convertDoc2Pdf(String inputFilePath, String outputFilePath);
    
    void convertXls2pdf(String inputFilePath, String outputFilePath);
    
    void convertXls2html(String inputFilePath, String outputFilePath);
}
