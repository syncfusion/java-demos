package rtl;

import java.io.*;
import com.syncfusion.docio.*;
import com.syncfusion.javahelper.system.io.*;
import com.syncfusion.javahelper.system.text.EncodingSupport;

public class RTL {
	public static void main(String[] args) throws Exception {
		// Creats a new document.
		WordDocument document = new WordDocument();
		// Adds a section & a paragraph in the empty document. 
		document.ensureMinimal();
		// Sets margins set up to the section. 
		document.getLastSection().getPageSetup().getMargins().setAll((float) 72);
		// Reads Arabic text from text file.
		StreamReaderSupport s = new StreamReaderSupport(
				(new File(getDataDir("Arabic.txt"))).toString(),
				EncodingSupport.getEncoding("ASCII"));
		String text = s.readToEnd();
		// Appends Arabic text to the document.
		document.getLastParagraph().getParagraphFormat().setBidi(true);
		// Appends text to the last paragraph.
		IWTextRange txtRange = document.getLastParagraph().appendText(text);
		// Sets character format to text.
		txtRange.getCharacterFormat().setBidi(true);
		// Sets the RTL text font size.
		txtRange.getCharacterFormat().setFontSizeBidi((float) 16);
		// Saves and closes the document.
		document.save("RTL.docx");
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
