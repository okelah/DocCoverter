package com.tencent.fm.convert;

import java.io.InputStream;
import java.io.OutputStream;


/**
 * Created by pengfeining on 2018/10/23 0023.
 */
public interface Convert {
    
    public void convertDoc2Html(InputStream inputStream, OutputStream outputStream);
    
    public void convertDoc2Pdf(InputStream inputStream, OutputStream outputStream);
}
