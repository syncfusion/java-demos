package bookmarknavigation;

import java.io.*;
import com.syncfusion.docio.*;
import com.syncfusion.javahelper.system.ObjectSupport;

public class BookmarkNavigation {
	public static void main(String[] args) throws Exception {
		// Create a new document.
		WordDocument document = new WordDocument();
		// Add section with one empty paragraph to the Word document.
		document.ensureMinimal();
		// Set the page margins.
		document.getLastSection().getPageSetup().getMargins().setAll(72f);
		// Append bookmark to the paragraph.
		WParagraph paragraph = document.getLastParagraph();
		paragraph.appendBookmarkStart("NorthwindDatabase");
		paragraph.appendText("Northwind database with normalization concept");
		paragraph.appendBookmarkEnd("NorthwindDatabase");
		// Open an existing template document with single section to get
		// Northwind.information.
		String path = getDataDir("Bookmark_Template.docx");
		WordDocument nwdInformation = new WordDocument(path);
		// Open an existing template document with multiple section to get Northwind
		// data.
		WordDocument templateDocument = new WordDocument(getDataDir("BkmkDocumentPart_Template.docx"));
		// Create a bookmark navigator which help us to navigate through the
		// bookmarks in the template document.
		BookmarksNavigator bk = new BookmarksNavigator(templateDocument);
		// Move to the NorthWind bookmark in template document.
		bk.moveToBookmark("NorthWind");
		// Get the bookmark content as WordDocumentPart.
		WordDocumentPart documentPart = bk.getContent();
		// Creating a bookmark navigator. Which help us to navigate through the
		// bookmarks in the Northwind information document.
		bk = new BookmarksNavigator(nwdInformation);
		// Move to the information bookmark.
		bk.moveToBookmark("Information");
		// Get the content of information bookmark.
		TextBodyPart bodyPart = bk.getBookmarkContent();
		// Create bookmark navigator which help us to navigate through the
		// bookmarks in the destination document.
		bk = new BookmarksNavigator(document);
		// Move to the NorthWind database in the destination document.
		bk.moveToBookmark("NorthwindDatabase");
		// Replace the bookmark content using word document parts.
		bk.replaceContent(documentPart);
		// Move to the Northwind_Information in the destination document.
		bk.moveToBookmark("Northwind_Information");
		// Replacing content of Northwind_Information bookmark.
		bk.replaceBookmarkContent(bodyPart);
		// Move to the text bookmark.
		bk.moveToBookmark("Text");
		// Delete the bookmark content.
		bk.deleteBookmarkContent(true);
		// Inserts text inside the bookmark. This will preserve the source formatting.
		bk.insertText("Northwind Database contains the following table:");

		// Append table.
		WTable table = new WTable(document);
		RowFormat tableFormat = table.getTableFormat();
		tableFormat.getBorders().setBorderType(BorderStyle.None);
		tableFormat.setIsAutoResized(true);
		table.resetCells(8, 2);
		table.getRows().get(0).setIsHeader(true);

		// Create table.
		IWParagraph paragraphs = createTable(table, 0, 0, "Suppliers");
		paragraphs = createTable(table, 0, 1, "1");
		paragraphs = createTable(table, 1, 0, "Customers");
		paragraphs = createTable(table, 1, 1, "1");
		paragraphs = createTable(table, 2, 0, "Employees");
		paragraphs = createTable(table, 2, 1, "3");
		paragraphs = createTable(table, 3, 0, "Products");
		paragraphs = createTable(table, 3, 1, "1");
		paragraphs = createTable(table, 4, 0, "Inventory");
		paragraphs = createTable(table, 4, 1, "2");
		paragraphs = createTable(table, 5, 0, "Shippers");
		paragraphs = createTable(table, 5, 1, "1");
		paragraphs = createTable(table, 6, 0, "PO Transactions");
		paragraphs = createTable(table, 6, 1, "3");
		paragraphs = createTable(table, 7, 0, "Sales Transactions");
		paragraphs = createTable(table, 7, 1, "7");
		bk.insertTable(table);
		// Move to image bookmark.
		bk.moveToBookmark("Image");
		// Deletes the bookmark content.
		bk.deleteBookmarkContent(true);

		// Inserting image to the bookmark.
		IWPicture pic = (WPicture) bk.insertParagraphItem(ParagraphItemType.Picture);
		pic.loadImage(new FileInputStream(getDataDir("Northwind.png")));
		// It reduces the image size because it doesnot fit in document page.
		pic.setWidthScale(50f);
		pic.setHeightScale(75f);
		bodyPart.close();
		documentPart.close();

		// Save and close the document.
		document.save("Bookmark Navigation.docx");
		document.close();
		System.out.println("Word document generated successfully.");
	}

	/**
	 * 
	 * Create table.
	 * 
	 * @param row    specifies the row index.
	 * @param column specifies the column index.
	 * @param text   specifies the text to be added.
	 * 
	 */

	private static IWParagraph createTable(WTable table, int row, int column, String text) throws Exception {
		// Append paragraph to the table cell.
		IWParagraph paragraph = table.get(row, column).addParagraph();
		// Add text to the paragraph.
		paragraph.appendText(text);
		// Applying character formats to text.
		WCharacterFormat characterFormat = paragraph.getBreakCharacterFormat();
		characterFormat.setFontName("Calibri");
		characterFormat.setFontSize(10);
		return paragraph;
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
