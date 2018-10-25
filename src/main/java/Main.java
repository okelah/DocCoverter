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
//        String testfile="my.docx";
        String testfile="chart.docx";
        doc4jPdfTest(testfile);
        doc4jPdfTest(testfile);
        asposePdfTest(testfile);
        asposePdfTest(testfile);
        
    }
    
    static void doc4jPdfTest(String filename) throws Exception {
        String pathInput = "E:\\ConvertTest\\"+filename;
        String pathOutput = "E:\\ConvertTest\\doc4j\\"+filename.substring(0,filename.indexOf("."))+".pdf";
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
        String pathInput = "E:\\ConvertTest\\"+filename;
        String pathOutput = "E:\\ConvertTest\\aspose\\"+filename.substring(0,filename.indexOf("."))+".pdf";
        // asposeConvert
        long asposeStartTime = System.currentTimeMillis();
        FileInputStream fileInputStream = new FileInputStream(new File(pathInput));
        FileOutputStream fileOutputStream = new FileOutputStream(new File(pathOutput));

        Convert asposeConvert = new AsposeConvert();
        asposeConvert.convertDoc2Pdf(fileInputStream, fileOutputStream);
        long asposeEndTime = System.currentTimeMillis();
        logger.info("aspose cost {} ms", asposeEndTime - asposeStartTime);
    }
}
