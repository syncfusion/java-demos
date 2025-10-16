package htmltoword;

import java.io.File;

import com.syncfusion.docio.*;
import com.syncfusion.javahelper.system.io.StreamReaderSupport;

public class HTMLToWord {
	public static void main(String[] args) throws Exception {
		// Creates a new Word document
		WordDocument document = new WordDocument();
		// Adds new section to the document
		IWSection section = document.addSection();
		// Set page size of the section
		section.getPageSetup().getMargins().setAll(72);
		// Adds new paragraph to the section
		IWParagraph para = section.addParagraph();
		StreamReaderSupport read = new StreamReaderSupport(getDataDir("Transitional.html"));
		document.setXHTMLValidateOption(XHTMLValidationType.Transitional);
		section.getBody().insertXHTML(read.readToEnd());
		// Save and close the document.
		document.save("HTML to Word.docx");
		document.close();
		System.out.println("Word document generated successfully.");
	}
	/**
	 * Get the file path
	 * 
	 * @param path specifies the file path
	 */
	public static String getDataDir(String path) {
        File dir = new File(System.getProperty("user.dir"));
		if(!(dir.toString().endsWith("samples")))
			dir = dir.getParentFile();
        dir = new File(dir, "resources");
        dir = new File(dir, path);
        if (dir.isDirectory() == false)
            dir.mkdir();
        return dir.toString();
    }
}
