package helloworld;

import java.io.File;
import java.io.FileInputStream;
import com.syncfusion.docio.*;
import com.syncfusion.javahelper.system.*;
import com.syncfusion.javahelper.system.drawing.*;

public class HelloWorld {
	public static void main(String[] args) throws Exception {
		// Creating a new document.
		WordDocument document = new WordDocument();
		// Adding a new section to the document.
		WSection section = (WSection) document.addSection();
		// Set Margin of the section
		section.getPageSetup().getMargins().setAll((float) 72);
		// Set page size of the section
		section.getPageSetup().setPageSize(new SizeFSupport(612, 792));
		// Create Paragraph styles
		WParagraphStyle style = createParagraphStyle(document);
		// Creates paragraph and apply formatting
		createParagraph(section);
		// Creates table and apply formatting
		createTable(section);
		// Save and close the document.
		document.save("Hello World.docx");
		document.close();
		System.out.println("Word document generated successfully.");
	}

	/**
	 * Creates paragraph and apply formatting
	 * 
	 * @param document Represents the word document
	 * 
	 */
	static WParagraphStyle createParagraphStyle(WordDocument document) throws Exception {

		// Applying the styles
		WParagraphStyle style = (WParagraphStyle) document.addParagraphStyle("Normal");
		WCharacterFormat characterFormat = style.getCharacterFormat();
		WParagraphFormat paragraphFormat = style.getParagraphFormat();

		appendCharacterFormatToText(style.getCharacterFormat(), 11f, "Calibri");
		paragraphFormat.setBeforeSpacing((float) 0);
		paragraphFormat.setAfterSpacing((float) 8);
		paragraphFormat.setLineSpacing(13.8f);
		style = (WParagraphStyle) document.addParagraphStyle("Heading 1");

		characterFormat = style.getCharacterFormat();
		paragraphFormat = style.getParagraphFormat();
		style.applyBaseStyle("Normal");

		appendCharacterFormatToText(style.getCharacterFormat(), 16f, "Calibri Light");
		characterFormat.setTextColor((ColorSupport.fromArgb(46, 116, 181)).clone());
		paragraphFormat.setBeforeSpacing((float) 12);
		paragraphFormat.setAfterSpacing((float) 0);
		paragraphFormat.setKeep(true);
		paragraphFormat.setKeepFollow(true);
		paragraphFormat.setOutlineLevel(OutlineLevel.Level1);
		return style;
	}

	/**
	 *
	 * Creates paragraph and apply formatting
	 * 
	 * @param section Represents the section of a Word document
	 */
	static IWParagraph createParagraph(WSection section) throws Exception {
		IWParagraph paragraph = section.getHeadersFooters().getHeader().addParagraph();

		// Append picture to the paragraph.
		WPicture picture = (WPicture) paragraph
				.appendPicture(new FileInputStream(getDataDir("AdventureCycle.jpg")));
		picture.setTextWrappingStyle(TextWrappingStyle.InFrontOfText);
		picture.setVerticalOrigin(VerticalOrigin.Margin);
		picture.setVerticalPosition((float) -45);
		picture.setHorizontalOrigin(HorizontalOrigin.Column);
		picture.setHorizontalPosition(263.5f);
		picture.setWidthScale((float) 20);
		picture.setHeightScale((float) 15);
		paragraph.applyStyle("Normal");
		paragraph.getParagraphFormat().setHorizontalAlignment(HorizontalAlignment.Left);

		// Add text to the paragraph.
		WTextRange textRange = (WTextRange) paragraph.appendText("Adventure Works Cycles");
		appendCharacterFormatToText(textRange.getCharacterFormat(), 12f, "Calibri");
		textRange.getCharacterFormat().setTextColor((ColorSupport.getRed()).clone());

		// Appends paragraph.
		paragraph = section.addParagraph();
		paragraph.applyStyle("Heading 1");
		paragraph.getParagraphFormat().setHorizontalAlignment(HorizontalAlignment.Center);

		// Add text to the paragraph.
		textRange = (WTextRange) paragraph.appendText("Adventure Works Cycles");
		appendCharacterFormatToText(textRange.getCharacterFormat(), 18f, "Calibri");

		// Appends paragraph.
		paragraph = section.addParagraph();
		paragraph.getParagraphFormat().setFirstLineIndent((float) 36);
		paragraph.getBreakCharacterFormat().setFontSize(12f);

		// Add text to the paragraph.
		textRange = (WTextRange) paragraph.appendText(
				"Adventure Works Cycles, the fictitious company on which the AdventureWorks sample databases are based, is a large, multinational manufacturing company. The company manufactures and sells metal and composite bicycles to North American, European and Asian commercial markets. While its base operation is located in Bothell, Washington with 290 employees, several regional sales teams are located throughout their market base.");
		textRange.getCharacterFormat().setFontSize(12f);

		// Appends paragraph.
		paragraph = section.addParagraph();
		paragraph.getParagraphFormat().setFirstLineIndent((float) 36);
		paragraph.getBreakCharacterFormat().setFontSize(12f);

		// Add text to the paragraph.
		textRange = (WTextRange) paragraph.appendText(
				"In 2000, Adventure Works Cycles bought a small manufacturing plant, Importadores Neptuno, located in Mexico. Importadores Neptuno manufactures several critical subcomponents for the Adventure Works Cycles product line. These subcomponents are shipped to the Bothell location for final product assembly. In 2001, Importadores Neptuno, became the sole manufacturer and distributor of the touring bicycle product group.");
		textRange.getCharacterFormat().setFontSize(12f);

		// Appends paragraph.
		paragraph = section.addParagraph();
		paragraph.applyStyle("Heading 1");
		paragraph.getParagraphFormat().setHorizontalAlignment(HorizontalAlignment.Left);

		// Add text to the paragraph.
		textRange = (WTextRange) paragraph.appendText("Product Overview");
		appendCharacterFormatToText(textRange.getCharacterFormat(), 16f, "Calibri");
		return paragraph;
	}

	/**
	 * Creates table and apply formatting
	 * 
	 * @param section Represents the section of a Word document
	 */
	static IWTable createTable(WSection section) throws Exception {

		// Appends table
		IWTable table = section.addTable();
		table.resetCells(3, 2);
		table.getTableFormat().getBorders().setBorderType(BorderStyle.None);
		table.getTableFormat().setIsAutoResized(true);

		// Appends paragraph.
		IWParagraph paragraph = table.get(0, 0).addParagraph();
		paragraph.getParagraphFormat().setAfterSpacing((float) 0);
		paragraph.getBreakCharacterFormat().setFontSize(12f);

		// Append picture to paragraph.
		IWPicture picture = (WPicture) paragraph
				.appendPicture(new FileInputStream(getDataDir("Mountain-200.jpg")));
		picture.setTextWrappingStyle(TextWrappingStyle.TopAndBottom);
		picture.setVerticalOrigin(VerticalOrigin.Paragraph);
		picture.setVerticalPosition(4.5f);
		picture.setHorizontalOrigin(HorizontalOrigin.Column);
		picture.setHorizontalPosition(-2.15f);
		picture.setWidthScale((float) 79);
		picture.setHeightScale((float) 79);

		// Appends paragraph.
		paragraph = table.get(0, 1).addParagraph();
		paragraph.applyStyle("Heading 1");
		appendParagraphFormatToParagraph(paragraph.getParagraphFormat(), (float) 0, 12f);
		paragraph.appendText("Mountain-200");

		// Appends paragraph.
		paragraph = table.get(0, 1).addParagraph();
		appendParagraphFormatToParagraph(paragraph.getParagraphFormat(), (float) 0, 12f);
		appendCharacterFormatToText(paragraph.getBreakCharacterFormat(), 12f, "Times New Roman");

		// Add text to the paragraph.
		WTextRange textRange = (WTextRange) paragraph.appendText("Product No: BK-M68B-38\r");
		appendCharacterFormatToText(textRange.getCharacterFormat(), 12f, "Times New Roman");

		// Add text to the paragraph.
		textRange = (WTextRange) paragraph.appendText("Size: 38\r");
		appendCharacterFormatToText(textRange.getCharacterFormat(), 12f, "Times New Roman");

		// Add text to the paragraph.
		textRange = (WTextRange) paragraph.appendText("Weight: 25\r");
		appendCharacterFormatToText(textRange.getCharacterFormat(), 12f, "Times New Roman");

		// Add text to the paragraph.
		textRange = (WTextRange) paragraph.appendText("Price: $2,294.99\r");
		appendCharacterFormatToText(textRange.getCharacterFormat(), 12f, "Times New Roman");

		// Appends paragraph.
		paragraph = table.get(0, 1).addParagraph();
		appendParagraphFormatToParagraph(paragraph.getParagraphFormat(), (float) 0, 12f);
		paragraph.getBreakCharacterFormat().setFontSize(12f);

		// Appends paragraph.
		paragraph = table.get(1, 0).addParagraph();
		paragraph.applyStyle("Heading 1");
		appendParagraphFormatToParagraph(paragraph.getParagraphFormat(), (float) 0, 12f);
		paragraph.appendText("Mountain-300 ");

		// Appends paragraph.
		paragraph = table.get(1, 0).addParagraph();
		appendParagraphFormatToParagraph(paragraph.getParagraphFormat(), (float) 0, 12f);
		appendCharacterFormatToText(paragraph.getBreakCharacterFormat(), 12f, "Times New Roman");

		// Add text to the paragraph.
		textRange = (WTextRange) paragraph.appendText("Product No: BK-M47B-38\r");
		appendCharacterFormatToText(textRange.getCharacterFormat(), 12f, "Times New Roman");

		// Add text to the paragraph.
		textRange = (WTextRange) paragraph.appendText("Size: 35\r");
		appendCharacterFormatToText(textRange.getCharacterFormat(), 12f, "Times New Roman");

		// Add text to the paragraph.
		textRange = (WTextRange) paragraph.appendText("Weight: 22\r");
		appendCharacterFormatToText(textRange.getCharacterFormat(), 12f, "Times New Roman");

		// Add text to the paragraph.
		textRange = (WTextRange) paragraph.appendText("Price: $1,079.99\r");
		appendCharacterFormatToText(textRange.getCharacterFormat(), 12f, "Times New Roman");

		// Appends paragraph.
		paragraph = table.get(1, 0).addParagraph();
		appendParagraphFormatToParagraph(paragraph.getParagraphFormat(), (float) 0, 12f);
		paragraph.getBreakCharacterFormat().setFontSize(12f);

		// Appends paragraph.
		paragraph = table.get(1, 1).addParagraph();
		paragraph.applyStyle("Heading 1");
		paragraph.getParagraphFormat().setLineSpacing(12f);

		// Append picture to the paragraph.
		picture = (WPicture) paragraph
				.appendPicture(new FileInputStream(getDataDir("Mountain-300.jpg")));
		picture.setTextWrappingStyle(TextWrappingStyle.TopAndBottom);
		picture.setVerticalOrigin(VerticalOrigin.Paragraph);
		picture.setVerticalPosition(8.2f);
		picture.setHorizontalOrigin(HorizontalOrigin.Column);
		picture.setHorizontalPosition(-14.95f);
		picture.setWidthScale((float) 75);
		picture.setHeightScale((float) 75);

		// Append paragraph.
		paragraph = table.get(2, 0).addParagraph();
		paragraph.applyStyle("Heading 1");
		paragraph.getParagraphFormat().setLineSpacing(12f);

		// Append picture to the paragraph.
		picture = (WPicture) paragraph
				.appendPicture(new FileInputStream(getDataDir("Road-550-W.jpg")));
		picture.setTextWrappingStyle(TextWrappingStyle.TopAndBottom);
		picture.setVerticalOrigin(VerticalOrigin.Paragraph);
		picture.setVerticalPosition(3.75f);
		picture.setHorizontalOrigin(HorizontalOrigin.Column);
		picture.setHorizontalPosition(-5f);
		picture.setWidthScale((float) 92);
		picture.setHeightScale((float) 92);

		// Append paragraph.
		paragraph = table.get(2, 1).addParagraph();
		paragraph.applyStyle("Heading 1");
		appendParagraphFormatToParagraph(paragraph.getParagraphFormat(), (float) 0, 12f);
		paragraph.appendText("Road-150 ");

		// Append paragraph.
		paragraph = table.get(2, 1).addParagraph();
		appendParagraphFormatToParagraph(paragraph.getParagraphFormat(), (float) 0, 12f);
		appendCharacterFormatToText(paragraph.getBreakCharacterFormat(), 12f, "Times New Roman");

		// Add text to the paragraph.
		textRange = (WTextRange) paragraph.appendText("Product No: BK-R93R-44\r");
		appendCharacterFormatToText(textRange.getCharacterFormat(), 12f, "Times New Roman");

		// Add text to the paragraph.
		textRange = (WTextRange) paragraph.appendText("Size: 44\r");
		appendCharacterFormatToText(textRange.getCharacterFormat(), 12f, "Times New Roman");

		// Add text to the paragraph.
		textRange = (WTextRange) paragraph.appendText("Weight: 14\r");
		appendCharacterFormatToText(textRange.getCharacterFormat(), 12f, "Times New Roman");

		// Add text to the paragraph.
		textRange = (WTextRange) paragraph.appendText("Price: $3,578.27\r");
		appendCharacterFormatToText(textRange.getCharacterFormat(), 12f, "Times New Roman");

		// Append paragraph.
		paragraph = table.get(2, 1).addParagraph();
		paragraph.applyStyle("Heading 1");
		paragraph.getParagraphFormat().setLineSpacing(12f);
		section.addParagraph();
		return table;
	}

	/**
	 * 
	 * @param Apply characterFormat
	 * @throws Exception
	 */
	static void appendCharacterFormatToText(WCharacterFormat characterFormat, float fontSize, String fontName)
			throws Exception {
		characterFormat.setFontSize(fontSize);
		characterFormat.setFontName(fontName);
	}

	/**
	 * 
	 * @param Apply paragraphFormat
	 * @throws Exception
	 */
	static void appendParagraphFormatToParagraph(WParagraphFormat paragraphFormat, float afterSpacingValue,
			float lineSpacingValue) throws Exception {
		paragraphFormat.setAfterSpacing(afterSpacingValue);
		paragraphFormat.setLineSpacing(lineSpacingValue);
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
