package footnotesandendnotes;

import java.io.*;
import com.syncfusion.docio.*;
import com.syncfusion.javahelper.system.drawing.*;

public class FootnotesandEndnotes {
	public static void main(String[] args) throws Exception {
		// A new document is created.
		WordDocument document = new WordDocument();
		// Creates footnotes at the bottom of the page.
		createFootNote(document);
		// Creates endnotes at the end of the section.
		createEndNote(document);
		// Save and close the document.
		document.save("Footnotes and Endnotes.docx");
		document.close();
		System.out.println("Word document generated successfully.");
	}

	/**
	 * Create footnotes at the bottom of the page.
	 * 
	 * @param document Represents a Word document instance.
	 *
	 */
	static void createFootNote(WordDocument document) throws Exception {
		// Add a new section to the document.
		IWSection section = document.addSection();
		section.getPageSetup().getMargins().setAll((float) 72);
		// Add a new paragraph to the section.
		IWParagraph paragraph = section.addParagraph();
		// Add text to the paragraph.
		IWTextRange textRange = paragraph.appendText("\t\t\t\t\tDemo for Footnote");
		// Applies character formats to text.
		appendCharacterFormatToText(textRange.getCharacterFormat(), ColorSupport.getBlack(), (float) 20, true);

		// Add paragraph to the section.
		section.addParagraph();
		section.addParagraph();
		paragraph = section.addParagraph();

		// Append footnote.
		WFootnote footnote = new WFootnote(document);
		footnote = paragraph.appendFootnote(FootnoteType.Footnote);
		footnote.getMarkerCharacterFormat().setSubSuperScript(SubSuperScript.SuperScript);

		// Insert text into the paragraph.
		paragraph.appendText("Google").getCharacterFormat().setBold(true);

		// Add paragraph to the section.
		paragraph = section.addParagraph();
		paragraph.getParagraphFormat().setHorizontalAlignment(HorizontalAlignment.Center);
		paragraph.appendPicture(new FileInputStream(getDataDir("google.png")));
		paragraph = footnote.getTextBody().addParagraph();

		// Insert text into the paragraph.
		paragraph.appendText(" Google is the most famous search engines in the Word ");

		// Add paragraph to the section.
		paragraph = section.addParagraph();
		paragraph = section.addParagraph();

		// Append footnote to the paragraph.
		footnote = paragraph.appendFootnote(FootnoteType.Footnote);
		footnote.getMarkerCharacterFormat().setSubSuperScript(SubSuperScript.SuperScript);

		// Insert text into the paragraph.
		paragraph.appendText("Yahoo").getCharacterFormat().setBold(true);

		// Add paragraph to the section.
		paragraph = section.addParagraph();
		paragraph.getParagraphFormat().setHorizontalAlignment(HorizontalAlignment.Center);
		paragraph.appendPicture(new FileInputStream(getDataDir("yahoo.gif")));
		paragraph = footnote.getTextBody().addParagraph();

		// Insert text into the paragraph.
		paragraph.appendText(
				" Yahoo experience makes it easier to discover the news and information that you care about most. ");

		// Add paragraph to the section.
		paragraph = section.addParagraph();
		paragraph = section.addParagraph();

		// Append footnote to the paragraph.
		footnote = paragraph.appendFootnote(FootnoteType.Footnote);
		footnote.getMarkerCharacterFormat().setSubSuperScript(SubSuperScript.SuperScript);

		// Insert text into the paragraph.
		paragraph.appendText("Northwind Traders").getCharacterFormat().setBold(true);

		// Add paragraph to the section.
		paragraph = section.addParagraph();
		paragraph.getParagraphFormat().setHorizontalAlignment(HorizontalAlignment.Center);
		paragraph.appendPicture(new FileInputStream(getDataDir("Northwind_logo.png")));
		paragraph = footnote.getTextBody().addParagraph();

		// Insert text into the paragraph.
		paragraph.appendText(
				" The Northwind sample database (Northwind.mdb) is included with all versions of Access. It provides data you can experiment with and database objects that demonstrate features you might want to implement in your own databases. ");
		// Sets number format for Footnote.
		document.setFootnoteNumberFormat(FootEndNoteNumberFormat.Arabic);
		// Sets Footnote position.
		document.setFootnotePosition(FootnotePosition.PrintAtBottomOfPage);
	}

	/**
	 * Create endnotes at the end of the section.
	 * 
	 * @param document Represents a Word document instance
	 */
	static void createEndNote(WordDocument document) throws Exception {
		// Add a new section to the document.
		IWSection section = document.addSection();
		// Add a new paragraph to the section.
		IWParagraph paragraph = section.addParagraph();
		// Add text to the paragraph.
		IWTextRange textRange = paragraph.appendText("\t\t\t\t\tDemo for Endnote");
		// Apply character formats to text.
		appendCharacterFormatToText(textRange.getCharacterFormat(), ColorSupport.getBlack(), (float) 20, true);

		// Add paragraph to the section.
		section.addParagraph();
		section.addParagraph();
		paragraph = section.addParagraph();

		// Append endnote to the paragraph.
		WFootnote footnote = new WFootnote(document);
		footnote = paragraph.appendFootnote(FootnoteType.Endnote);
		footnote.getMarkerCharacterFormat().setSubSuperScript(SubSuperScript.SuperScript);

		// Insert text into the paragraph.
		paragraph.appendText("Google").getCharacterFormat().setBold(true);

		// Add paragraph to the section.
		paragraph = section.addParagraph();
		paragraph.getParagraphFormat().setHorizontalAlignment(HorizontalAlignment.Center);
		paragraph.appendPicture(new FileInputStream(getDataDir("google.png")));
		paragraph = footnote.getTextBody().addParagraph();

		// Insert text into the paragraph.
		paragraph.appendText(" Google is the most famous search engines in the Word ");
		section = document.addSection();
		section.setBreakCode(SectionBreakCode.NoBreak);

		// Add paragraph to the section.
		paragraph = section.addParagraph();
		paragraph = section.addParagraph();

		// Append endnote to the paragraph.
		footnote = paragraph.appendFootnote(FootnoteType.Endnote);
		footnote.getMarkerCharacterFormat().setSubSuperScript(SubSuperScript.SuperScript);

		// Insert text into the paragraph.
		paragraph.appendText("Yahoo").getCharacterFormat().setBold(true);

		// Add paragraph to the section.
		paragraph = section.addParagraph();
		paragraph.getParagraphFormat().setHorizontalAlignment(HorizontalAlignment.Center);
		paragraph.appendPicture(new FileInputStream(getDataDir("yahoo.gif")));
		paragraph = footnote.getTextBody().addParagraph();

		// Insert text into the paragraph.
		paragraph.appendText(
				" Yahoo experience makes it easier to discover the news and information that you care about most. ");

		// Add paragraph to the section.
		section = document.addSection();
		section.setBreakCode(SectionBreakCode.NoBreak);

		// Add paragraph to the section.
		paragraph = section.addParagraph();
		paragraph = section.addParagraph();
		footnote = paragraph.appendFootnote(FootnoteType.Endnote);
		footnote.getMarkerCharacterFormat().setSubSuperScript(SubSuperScript.SuperScript);

		// Insert text into the paragraph.
		paragraph.appendText("Northwind Traders").getCharacterFormat().setBold(true);
		paragraph = section.addParagraph();
		paragraph.getParagraphFormat().setHorizontalAlignment(HorizontalAlignment.Center);
		paragraph.appendPicture(new FileInputStream(getDataDir("Northwind_logo.png")));
		paragraph = footnote.getTextBody().addParagraph();

		// Insert text into the paragraph.
		paragraph.appendText(
				" The Northwind sample database (Northwind.mdb) is included with all versions of Access. It provides data you can experiment with and database objects that demonstrate features you might want to implement in your own databases ");
		// Sets number format for endnote.
		document.setEndnoteNumberFormat(FootEndNoteNumberFormat.LowerCaseRoman);
		// Sets restart index for endnote.
		document.setRestartIndexForEndnote(EndnoteRestartIndex.DoNotRestart);
		// Sets endnote position.
		document.setEndnotePosition(EndnotePosition.DisplayEndOfSection);
	}

	/**
	 * 
	 * @param Apply characterFormat
	 * @throws Exception
	 */
	static void appendCharacterFormatToText(WCharacterFormat characterFormat, ColorSupport color, float fontSize,
			boolean value) throws Exception {
		characterFormat.setTextColor(color);
		characterFormat.setBold(value);
		characterFormat.setFontSize(fontSize);
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
