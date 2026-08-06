package headerandfooter;

import java.io.*;
import java.nio.charset.Charset;
import com.syncfusion.docio.*;
import com.syncfusion.javahelper.system.*;
import com.syncfusion.javahelper.system.io.*;

public class HeaderandFooter {
	public static void main(String[] args) throws Exception {
		// Creates a new document.
		WordDocument doc = new WordDocument();
		// Adds a new section to the document.
		IWSection section1 = doc.addSection();
		// Sets the header/footer setup.
		section1.getPageSetup().setDifferentFirstPage(true);
		// Inserts Header Footer to first page
		insertFirstPageHeaderFooter(doc, section1, PageNumberStyle.Arabic);
		// Inserts Header Footer to all pages
		insertPageHeaderFooter(doc, section1, PageNumberStyle.Arabic);
		// Adds text to the document body section.
		IWParagraph par;
		par = section1.addParagraph();
		// Inserts Text into the word Document.
		StreamReaderSupport reader = new StreamReaderSupport(
				new MemoryStreamSupport(new FileInputStream(getDataDir("WinFAQ.txt"))), Charset.forName("ASCII"));
		String text = reader.readToEnd();
		// Appends text to the paragraph.
		par.appendText(text);
		// Saves and closes the document.
		doc.save("Header and Footer.docx");
		doc.close();
		System.out.println("Word document generated successfully.");
	}

	/**
	 * Inserts Header Footer to first page
	 * 
	 * @param doc      Represent the Word document instance
	 * @param section  Represents a section.
	 * @param numStyle Represents number style for a page.
	 */

	private static void insertFirstPageHeaderFooter(WordDocument doc, IWSection section, PageNumberStyle numStyle)
			throws Exception {
		// Adds a new paragraph for header to the document.
		IWParagraph headerPar = new WParagraph(doc);
		// Adds a new table to the header.
		IWTable table = section.getHeadersFooters().getFirstPageHeader().addTable();
		RowFormat format = new RowFormat();
		// Sets cleared table border style.
		format.getBorders().setBorderType(BorderStyle.Cleared);
		// Inserts table with a row and two columns.
		table.resetCells(1, 2, format, (float) 265);
		// Adds paragraph to first row of first cell.
		headerPar = (WParagraph)table.get(0, 0).addParagraph();
		// Appends picture to the paragraph.
		headerPar.appendPicture(new FileInputStream(getDataDir("Northwind_logo.png")));
		// Sets width and height of picture. 
		((WPicture)headerPar.getItems().get(0)).setWidth(232.5f);
		((WPicture)headerPar.getItems().get(0)).setHeight(54.75f);
		headerPar = (WParagraph)table.get(0, 1).addParagraph();
		// Inserts text to the table second cell.
		IWTextRange txt = headerPar.appendText(
				"Company Headquarters,\n2501 Aerial Center Parkway,\nSuite 110, Morrisville, NC 27560,\nTEL 1-888-936-8638.");
		WCharacterFormat characterFormat = txt.getCharacterFormat();
		characterFormat.setFontSize((float) 12);
		characterFormat.setCharacterSpacing(1.7f);
		headerPar.getParagraphFormat().setHorizontalAlignment(HorizontalAlignment.Right);
		// Adds a new paragraph to the header with address text.
		headerPar = new WParagraph(doc);
		headerPar.getParagraphFormat().setHorizontalAlignment(HorizontalAlignment.Center);
		txt = headerPar.appendText("\nFirst Page Header");
		txt.getCharacterFormat().setCharacterSpacing(1.7f);
		section.getHeadersFooters().getFirstPageHeader().getParagraphs().add(headerPar);
		// Adds a footer paragraph text to the document.
		WParagraph footerPar = new WParagraph(doc);
		footerPar.getParagraphFormat().getTabs().addTab((float) 523f, TabJustification.Right, TabLeader.NoLeader);
		footerPar.appendText("Copyright Syncfusion Inc. 2001 - 2012");
		// Adds page and Number of pages field to the document.
		footerPar.appendText("\tFirst Page ");
		// Appends page field to the footer paragraph.
		footerPar.appendField("Page", FieldType.FieldPage);
		section.getHeadersFooters().getFirstPageFooter().getParagraphs().add(footerPar);
		// Sets page set up to the section.
		WPageSetup pageSetup = section.getPageSetup();
		pageSetup.setRestartPageNumbering(true);
		pageSetup.setPageStartingNumber(1);
		pageSetup.setPageNumberStyle(numStyle);
	}

	/**
	 * Inserts Header Footer to all pages
	 * 
	 * @param doc      Represent the Word document instance
	 * @param section1 Represents a section.
	 * @param numStyle Represents number style for a page.
	 */
	private static void insertPageHeaderFooter(WordDocument doc, IWSection section1, PageNumberStyle numStyle)
			throws Exception {
		// Adds a new paragraph for header to the document.
		IWParagraph headerPar = new WParagraph(doc);
		// Adds a new table to the header
		IWTable table = section1.getHeadersFooters().getHeader().addTable();
		RowFormat format = new RowFormat();
		// Sets Single table border style.
		format.getBorders().setBorderType(BorderStyle.Single);
		// Inserts table with a row and two columns.
		table.resetCells(1, 2, format, (float) 265);
		// Inserts logo image to the table first cell.
		headerPar = (WParagraph)table.get(0, 0).addParagraph();
		// Appends picture to the header paragraph.
		headerPar.appendPicture(new FileInputStream(getDataDir("Northwind_logo.png")));
		((WPicture)headerPar.getItems().get(0)).setWidth(232.5f);
		((WPicture)headerPar.getItems().get(0)).setHeight(54.75f);
		headerPar = (WParagraph) ObjectSupport.instanceOf(table.get(0, 1).addParagraph(), WParagraph.class);
		// Inserts text to the table second cell.
		IWTextRange txt = headerPar.appendText(
				"Company Headquarters,\n2501 Aerial Center Parkway,\nSuite 110, Morrisville, NC 27560,\nTEL 1-888-936-8638.");
		WCharacterFormat characterFormat = txt.getCharacterFormat();
		characterFormat.setFontSize((float) 12);
		characterFormat.setCharacterSpacing(1.7f);
		headerPar.getParagraphFormat().setHorizontalAlignment(HorizontalAlignment.Right);
		// Add a footer paragraph text to the document.
		WParagraph footerPar = new WParagraph(doc);
		footerPar.getParagraphFormat().getTabs().addTab((float) 523f, TabJustification.Right, TabLeader.NoLeader);
		// Add page and Number of pages field to the document.
		footerPar.appendText("Copyright Syncfusion Inc. 2001 - 2012");
		footerPar.appendText("\tPage ");
		IWField ff = footerPar.appendField("Page", FieldType.FieldPage);
		// Page Number Settings
		section1.getHeadersFooters().getFooter().getParagraphs().add(footerPar);
		// Sets page set up to the section.
		WPageSetup pageSetup = section1.getPageSetup();
		pageSetup.setRestartPageNumbering(true);
		pageSetup.setPageStartingNumber(1);
		pageSetup.setPageNumberStyle(numStyle);
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
