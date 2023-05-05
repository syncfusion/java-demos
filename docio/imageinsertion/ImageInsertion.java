package imageinsertion;

import java.io.*;
import com.syncfusion.docio.*;

public class ImageInsertion {
	public static void main(String[] args) throws Exception {
		// Creates a new document.
		WordDocument document = new WordDocument();
		// Adds a new section to the document.
		IWSection section = document.addSection();
		section.getPageSetup().getMargins().setAll((float) 72);
		// Adds a paragraph to the section
		IWParagraph paragraph = section.addParagraph();
		// Appends text to the paragraph.
		paragraph.appendText("This sample demonstrates how to insert Vector and Scalar images inside a document.");
		// Inserts .gif Image
		insertItemsToParagraph(paragraph, section, getDataDir("yahoo.gif"));
		applyFormattingForCaption(document.getLastParagraph());
		// Inserts .bmp Image.
		insertItemsToParagraph(paragraph, section, getDataDir("Reports.bmp"));
		applyFormattingForCaption(document.getLastParagraph());
		// Inserts .png image.
		insertItemsToParagraph(paragraph, section, getDataDir("google.PNG"));
		applyFormattingForCaption(document.getLastParagraph());
		// Inserts .tif image.
		insertItemsToParagraph(paragraph, section, getDataDir("Square.tif"));
		applyFormattingForCaption(document.getLastParagraph());
		// Adds paragraph to the section.
		paragraph = section.addParagraph();
		paragraph.getParagraphFormat().setHorizontalAlignment(HorizontalAlignment.Center);
		WPicture mImage = (WPicture) paragraph.appendPicture(new FileInputStream(getDataDir("Ess chart.emf")));
		// Sets width and height to the picture.
		mImage.setHeightScale(50f);
		mImage.setWidthScale(50f);
		// Sets caption format to picture.
		mImage.addCaption("Figure", CaptionNumberingFormat.Roman, CaptionPosition.AfterImage);
		applyFormattingForCaption(document.getLastParagraph());
		document.updateDocumentFields();
		// Saves and closes the document.
		document.save("Image Insertion.docx");
		document.close();
		System.out.println("Word document generated successfully.");

	}

	/**
	 * Inserts the paragraph item in paragraph.
	 * 
	 * @param paragraph
	 * @param section
	 * @param text
	 * @param alignment
	 * @param imagePath
	 * @throws Exception
	 */
	static void insertItemsToParagraph(IWParagraph paragraph, IWSection section, String imagePath) throws Exception {
		// Adds a new paragraph.
		paragraph = section.addParagraph();
		// Applies paragraph format.
		paragraph.getParagraphFormat().setHorizontalAlignment(HorizontalAlignment.Center);
		// Append picture to paragraph.
		WPicture mImage = (WPicture) paragraph.appendPicture(new FileInputStream(imagePath));
		// Adding Image caption
		mImage.addCaption("Figure", CaptionNumberingFormat.Roman, CaptionPosition.AfterImage);

	}

	/**
	 * Apply formatting for image caption paragraph
	 * 
	 * @param paragraph
	 * @throws Exception
	 */
	static void applyFormattingForCaption(WParagraph paragraph) throws Exception {
		// Align the caption
		paragraph.getParagraphFormat().setHorizontalAlignment(HorizontalAlignment.Center);
		// Sets after spacing
		paragraph.getParagraphFormat().setAfterSpacing(1.5f);
		// Sets before spacing
		paragraph.getParagraphFormat().setBeforeSpacing(1.5f);
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
