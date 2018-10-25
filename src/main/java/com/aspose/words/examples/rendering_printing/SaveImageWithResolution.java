package com.aspose.words.examples.rendering_printing;

import com.aspose.words.Document;
import com.aspose.words.ImageSaveOptions;
import com.aspose.words.SaveFormat;
import com.aspose.words.examples.Utils;

/**
 * Created by Home on 5/29/2017.
 */
public class SaveImageWithResolution {
    public static void main(String[] args) throws Exception {
        //ExStart:SetHorizontalAndVerticalImageResolution
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(SaveImageWithResolution.class);
        Document doc = new Document(dataDir + "TestFile.doc");

        //Renders a page of a Word document into a PNG image at a specific horizontal and vertical resolution.
        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.PNG);
        options.setHorizontalResolution(300);
        options.setVerticalResolution(300);
        options.setPageCount(1);

        doc.save(dataDir + "Rendering.SaveToImageResolution Out.png", options);
        //ExEnd:SetHorizontalAndVerticalImageResolution
    }
}
