package encrypt;

import com.syncfusion.docio.*;

public class Encrypt {
	public static void main(String[] args) throws Exception {
		// Initialize Word document.
		WordDocument document = new WordDocument();
		// Ensure Minimum.
        document.ensureMinimal();
        // Get last section of the document.
        IWSection section = document.getLastSection();
        // Add a paragraph to the section.
        IWParagraph paragraph = section.addParagraph();
        // Append text to the paragraph.
        IWTextRange text = paragraph.appendText("This document was encrypted with password");
        text.getCharacterFormat().setFontSize(16f);
        text.getCharacterFormat().setFontName("Bitstream Vera Serif");     
        // Encrypt the document by giving password
        document.encryptDocument("syncfusion");
        // Save and close the document.
     	document.save("Encrypt.docx");
     	document.close();
  		System.out.println("Word document generated successfully.");
	}

}
