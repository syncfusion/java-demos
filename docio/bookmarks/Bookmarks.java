package bookmarks;

import com.syncfusion.docio.*;
import com.syncfusion.javahelper.system.drawing.*;

public class Bookmarks {
	public static void main(String[] args) throws Exception {
		// Creating a new document.
		WordDocument document = new WordDocument();
		// Adding a section to the document.
		IWSection section = document.addSection();
		section.getPageSetup().getMargins().setAll((float) 72);
		// Adding a new paragraph to the section.
		IWParagraph paragraph = section.addParagraph();
		// Adding text to the paragraph.
		WTextRange textRange = (WTextRange) paragraph
				.appendText("This document demonstrates Essential DocIO's Bookmark functionality.");
		appendCharacterFormatToText(textRange.getCharacterFormat(), 14f);
		// Adding paragraph to the section
		section.addParagraph();
		paragraph = section.addParagraph();
		// Adding text to the paragraph.
		textRange = (WTextRange) paragraph.appendText("1. Inserting Bookmark Text");
		appendCharacterFormatToText(textRange.getCharacterFormat(), 12f);
		section.addParagraph();
		paragraph = section.addParagraph();
		// BookmarkStart.
		paragraph.appendBookmarkStart("Bookmark");
		// Write bookmark
		paragraph.appendText("Bookmark Text");
		paragraph.appendComment("This is a simple bookmark");
		// BookmarkEnd.
		paragraph.appendBookmarkEnd("Bookmark");
		// Adding paragraph to the section.
		section.addParagraph();
		paragraph = section.addParagraph();
		// Indicating hidden bookmark text start.
		paragraph.appendBookmarkStart("_HiddenText");
		// Writing bookmark text
		paragraph.appendText("2. Hidden Bookmark Text").getCharacterFormat()
				.setFont(new FontSupport("Comic Sans MS", 10));
		// Indicating hidden bookmark text end.
		paragraph.appendBookmarkEnd("_HiddenText");
		paragraph.appendComment("This is hidden bookmark");
		section.addParagraph();
		paragraph = section.addParagraph();
		// Writing nested bookmarks
		textRange = (WTextRange) paragraph.appendText("3. Nested Bookmarks");
		appendCharacterFormatToText(textRange.getCharacterFormat(), 12f);
		// Append paragraph to the section.
		section.addParagraph();
		paragraph = section.addParagraph();
		paragraph.appendBookmarkStart("Main");
		paragraph.appendText(" Main data ");
		paragraph.appendBookmarkStart("Nested");
		paragraph.appendText(" Nested data ");
		paragraph.appendComment("This is a nested bookmark");
		paragraph.appendBookmarkStart("NestedNested");
		paragraph.appendText(" Nested Nested ");
		paragraph.appendBookmarkEnd("NestedNested");
		paragraph.appendText(" data Nested ");
		paragraph.appendBookmarkEnd("Nested");
		paragraph.appendText(" Data Main ");
		paragraph.appendBookmarkEnd("Main");
		document.save("Bookmarks.docx");
		// Save and close the document
		document.close();
		System.out.println("Word document generated successfully.");
	}

	/**
	 * 
	 * @param Apply characterFormat
	 * @throws Exception
	 */
	static void appendCharacterFormatToText(WCharacterFormat characterFormat, float fontSize) throws Exception {
		characterFormat.setFontSize(fontSize);
	}
}
