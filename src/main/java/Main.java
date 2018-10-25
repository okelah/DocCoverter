import com.aspose.words.Document;
import com.aspose.words.HtmlMetafileFormat;
import com.aspose.words.HtmlSaveOptions;
import com.tencent.fm.convert.Convert;
import com.tencent.fm.convert.Impl.AsposeConvert;
import com.tencent.fm.convert.Impl.Doc4jConvert;
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

//        String testfile="mac_edit.docx";
//        String testfile = "mac-edit.docx";
//        String testfile = "chart.docx";
//        doc4jHtmlTest(testfile);
//        doc4jHtmlTest(testfile);
//        asposeHtmlTest(testfile);
//        asposeHtmlTest(testfile);
        saveHtmlWithMetafileFormat("E:\\ConvertTest\\");
//        doc4jPdfTest(testfile);
//        doc4jPdfTest(testfile);
//        asposePdfTest(testfile);
//        asposePdfTest(testfile);

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
//        String pathInput = "E:\\ConvertTest\\" + filename;
//        String pathOutput = "E:\\ConvertTest\\doc4j\\" + filename.substring(0, filename.indexOf(".")) + ".pdf";
        String pathInput = "/Users/null/Documents/SelfStudy/" + filename;
        String pathOutput = "/Users/null/Documents/SelfStudy/" + filename.substring(0, filename.indexOf(".")) + ".html";

        long doc4jStartTime = System.currentTimeMillis();
        // doc4jConvert
        FileInputStream fileInputStream = new FileInputStream(new File(pathInput));
        FileOutputStream fileOutputStream = new FileOutputStream(new File(pathOutput));
        Convert doc4jConvert = new Doc4jConvert();
        doc4jConvert.convertDoc2Pdf(fileInputStream, fileOutputStream);
        long doc4jEndTime = System.currentTimeMillis();
        logger.info("doc4j cost {} ms", doc4jEndTime - doc4jStartTime);
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
        asposeConvert.convertDoc2Pdf(fileInputStream, fileOutputStream);
        long asposeEndTime = System.currentTimeMillis();
        logger.info("aspose cost {} ms", asposeEndTime - asposeStartTime);
    }

    static void doc4jHtmlTest(String filename) throws Exception {
//        String pathInput = "E:\\ConvertTest\\" + filename;
//        String pathOutput = "E:\\ConvertTest\\doc4j\\" + filename.substring(0, filename.indexOf(".")) + ".html";
        String pathInput = "/Users/null/Documents/SelfStudy/" + filename;
        String pathOutput = "/Users/null/Documents/SelfStudy/" + filename.substring(0, filename.indexOf(".")) + ".html";

        long doc4jStartTime = System.currentTimeMillis();
        FileInputStream fileInputStream = new FileInputStream(new File(pathInput));
        FileOutputStream fileOutputStream = new FileOutputStream(new File(pathOutput));
        Convert convert = new Doc4jConvert();
        convert.convertDoc2Html(fileInputStream, fileOutputStream);
        long doc4jEndTime = System.currentTimeMillis();
        logger.info("aspose cost {} ms", doc4jStartTime - doc4jEndTime);
    }

    static void asposeHtmlTest(String filename) throws Exception {
//        String pathInput = "E:\\ConvertTest\\" + filename;
//        String pathOutput = "E:\\ConvertTest\\doc4j\\" + filename.substring(0, filename.indexOf(".")) + ".html";
        String pathInput = "/Users/null/Documents/SelfStudy/" + filename;
        String pathOutput = "/Users/null/Documents/SelfStudy/" + filename.substring(0, filename.indexOf(".")) + ".html";

        long doc4jStartTime = System.currentTimeMillis();
        FileInputStream fileInputStream = new FileInputStream(new File(pathInput));
        FileOutputStream fileOutputStream = new FileOutputStream(new File(pathOutput));
        Convert convert = new AsposeConvert();
        convert.convertDoc2Html(fileInputStream, fileOutputStream);
        long doc4jEndTime = System.currentTimeMillis();
        logger.info("aspose cost {} ms", doc4jStartTime - doc4jEndTime);
    }
}
