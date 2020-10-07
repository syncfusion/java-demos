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
		paragraph.appendText(
				"This Document demonstrates Essential DocIO's support for inserting Vector and Scalar Images inside a document.");
		// Inserts .gif Image
		insertItemsToParagraph(paragraph,section,"Essential XlsIO [.gif Image]",HorizontalAlignment.Center,getDataDir("XlsIO.gif"));
		// Inserts .bmp Image.
		insertItemsToParagraph(paragraph,section,"Essential Grid [.bmp Image]",HorizontalAlignment.Center,getDataDir("Ess Grid.bmp"));
		// Inserts .png image.
		insertItemsToParagraph(paragraph,section,"Essential Tools [.png Image]",HorizontalAlignment.Center,getDataDir("Ess Tools.PNG"));
		// Inserts .tif image.
		insertItemsToParagraph(paragraph,section,"Essential PDF [.tif Image]",HorizontalAlignment.Center,getDataDir("Ess PDF.tif"));
		// Adds paragraph to the section.
		paragraph = section.addParagraph();
		paragraph.getParagraphFormat().setHorizontalAlignment(HorizontalAlignment.Center);
		WPicture mImage = (WPicture) paragraph
				.appendPicture(new FileInputStream(getDataDir("Ess chart.emf")));
		// Sets width and height to the picture.
		mImage.setHeightScale(50f);
		mImage.setWidthScale(50f);
		// Sets caption format to picture.
		mImage.addCaption("Chart Vector Image", CaptionNumberingFormat.Roman, CaptionPosition.AboveImage);
		// Saves and closes the document.
		document.save("Image Insertion.docx");
		document.close();
		System.out.println("Word document generated successfully.");

	}
	/**
	 * Inserts the paragraph item in paragraph.
	 * @param paragraph
	 * @param section
	 * @param text
	 * @param alignment
	 * @param imagePath
	 * @throws Exception
	 */
	static void insertItemsToParagraph(IWParagraph paragraph, IWSection section ,String text, HorizontalAlignment alignment, String imagePath) throws Exception{
		// Adds a new paragraph.
	    paragraph = section.addParagraph();
	    // Append text to paragraph.
		paragraph.appendText(text);
		paragraph = section.addParagraph();
		// Applies paragraph format.
		paragraph.getParagraphFormat().setHorizontalAlignment(alignment);
		// Append picture to paragraph.
		paragraph.appendPicture(new FileInputStream(imagePath));
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
