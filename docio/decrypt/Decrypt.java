package decrypt;

import java.io.File;

import com.syncfusion.docio.*;

public class Decrypt {
	public static void main(String[] args) throws Exception {
		// Initialize Word document.
		WordDocument document = new WordDocument();
		document = new WordDocument();
		// Open an existing template document.
		document.open(getDataDir("Security Settings.docx"), FormatType.Docx, "syncfusion");
		// Get last section of the document.
		IWSection section = document.getLastSection();
		// Add a paragraph to the section.
		IWParagraph paragraph = section.addParagraph();
		// Append text to the paragraph.
		IWTextRange text = paragraph.appendText("\nDemo For Document Decryption with Essential DocIO");
		text.getCharacterFormat().setFontSize(16f);
		text.getCharacterFormat().setFontName("Bitstream Vera Serif");
		// Append text to the paragraph.
		text = paragraph.appendText("\nThis document is Decrypted");
		text.getCharacterFormat().setFontSize(16f);
		text.getCharacterFormat().setFontName("Bitstream Vera Serif");
		// Save and close the document.
		document.save("Decrypt.docx");
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
		if (!(dir.toString().endsWith("samples")))
			dir = dir.getParentFile();
		dir = new File(dir, "resources");
		dir = new File(dir, path);
		if (dir.isDirectory() == false)
			dir.mkdir();
		return dir.toString();
	}

}
