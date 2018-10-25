package com.tencent.fm.convert.Impl;

import com.aspose.words.*;
import com.tencent.fm.convert.Convert;

import java.io.InputStream;
import java.io.OutputStream;

/**
 * Created by pengfeining on 2018/10/25 0025.
 */
public class AsposeConvert implements Convert {
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
     * @param obj
     *            字符串对象
     * @return
     */
    protected static boolean notNull(String obj) {
        if (obj != null && !obj.equals("") && !obj.equals("undefined")
                && !obj.trim().equals("") && obj.trim().length() > 0) {
            return true;
        }
        return false;
    }


    public void convertDoc2Html(InputStream inputStream, OutputStream outputStream) {
        try {
            // 验证License
            if (!getLicense()) {
                return;
            }
            Document doc = new Document(inputStream);
            HtmlSaveOptions hso = new HtmlSaveOptions();
            hso.setExportRoundtripInformation(true);
            HtmlSaveOptions options = new HtmlSaveOptions();
            options.setMetafileFormat(HtmlMetafileFormat.EMF_OR_WMF);

            doc.save(outputStream, hso);//全面支持DOC, DOCX, OOXML, RTF HTML, OpenDocument, PDF, EPUB, XPS, SWF 相互转换
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public void convertDoc2Pdf(InputStream inputStream, OutputStream outputStream) {
        try {
            // 验证License
            if (!getLicense()) {
                return;
            }
                Document doc = new Document(inputStream);
                doc.save(outputStream, SaveFormat.PDF);//全面支持DOC, DOCX, OOXML, RTF HTML, OpenDocument, PDF, EPUB, XPS, SWF 相互转换
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
