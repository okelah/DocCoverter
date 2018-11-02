package com.tencent.fm.convert;

import com.tencent.fm.convert.bean.SourceFile;
import com.tencent.fm.convert.bean.SourceFileType;
import com.tencent.fm.convert.bean.TargetFile;
import com.tencent.fm.convert.bean.TargetFileType;

/**
 * Created by pengfeining on 2018/11/2 0002.
 */
public class FileConvertFactoryImpl implements FileConvertFactory {
    @Override
    public Convert create(SourceFile sourceFile, TargetFile targetFile) {
        if(sourceFile.getSourceFileType()== SourceFileType.DOC&&targetFile.getTargetFileType()== TargetFileType.HTML){//doc2html

        }else{

        }
        return null;
    }
}
