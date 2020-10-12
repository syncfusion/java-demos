package gettingstarted;

import com.syncfusion.docio.FormatType;
import java.io.FileInputStream;
import java.io.File;

import com.syncfusion.docio.*;

public class GettingStarted {

	public static void main(String[] args) throws Exception {
		//Create an instance of WordDocument Instance (Empty Word Document).
		WordDocument document = new WordDocument();
		//Add a new section into the Word document.
		IWSection section = document.addSection();
		//Specifies the page margins. 
		section.getPageSetup().getMargins().setAll(50f);
		//Add a new simple paragraph into the section.
		IWParagraph firstParagraph = section.addParagraph();
		//Set the paragraph's horizontal alignment as justify.
		firstParagraph.getParagraphFormat().setHorizontalAlignment(HorizontalAlignment.Justify);
		//Add a text range into the paragraph.
		IWTextRange firstTextRange = firstParagraph.appendText("AdventureWorks Cycles,");
		//set the font formatting of the text range.
		firstTextRange.getCharacterFormat().setBoldBidi(true);
		firstTextRange.getCharacterFormat().setFontName("Calibri");
		firstTextRange.getCharacterFormat().setFontSize(14) ;
		//Add another text range into the paragraph.
		IWTextRange secondTextRange = firstParagraph.appendText(" the fictitious company on which the AdventureWorks sample databases are based, is a large, multinational manufacturing company.");
		//set the font formatting of the text range.
		secondTextRange.getCharacterFormat().setFontName("Calibri");
		secondTextRange.getCharacterFormat().setFontSize(11);
		//Add another paragraph and aligns it as a center.
		IWParagraph paragraph = section.addParagraph();
		paragraph.getParagraphFormat().setHorizontalAlignment(HorizontalAlignment.Center);
		//Set after spacing for paragraph.
		paragraph.getParagraphFormat().setAfterSpacing(6);
		//Add a picture into the paragraph.
		IWPicture picture = paragraph.appendPicture(new FileInputStream(getDataDir("Image.png")));
		//Specify the size of the picture.
		picture.setHeight(86);
		picture.setWidth(81);
		IWTable table = section.addTable();
		//Create the specified number of rows and columns.
		table.resetCells(2,2);
		//Access the instance of the cell (first row, first cell).
		WTableCell firstCell = table.getRows().get(0).getCells().get(0);
		//Specifies the width of the cell.
		firstCell.setWidth(150);
		//Add a paragraph into the cell; a cell must have atleast 1 paragraph.
		paragraph=firstCell.addParagraph();		
		IWTextRange textRange = paragraph.appendText("Profile picture");
		textRange.getCharacterFormat().setBold(true);
		//Access the instance of cell (first row, second cell).
		WTableCell secondCell = table.getRows().get(0).getCells().get(1);
		secondCell.setWidth(330);
		paragraph=secondCell.addParagraph();
		//Add text to the paragraph. 
		textRange=paragraph.appendText("Description");
		textRange.getCharacterFormat().setBold(true);
		firstCell=table.getRows().get(1).getCells().get(0);
		firstCell.setWidth(150);
		//Add image to  the paragraph.
		paragraph=firstCell.addParagraph();
		//Set after spacing for paragraph.
		paragraph.getParagraphFormat().setAfterSpacing(6);
		IWPicture profilePicture = paragraph.appendPicture(new FileInputStream(getDataDir("Image.png")));
		//Set the height and width for the image.
		profilePicture.setHeight(98);
		profilePicture.setWidth(95);
		//Access the instance of cell (second row, second cell) and adds text.
		secondCell=table.getRows().get(1).getCells().get(1);
		secondCell.setWidth(330);
		paragraph=secondCell.addParagraph();
		textRange=paragraph.appendText("AdventureWorks Cycles, the fictitious company on which the AdventureWorks sample databases are based, is a large, multinational manufacturing company.");
		paragraph=section.addParagraph();
		//Set before spacing for paragraph.
		paragraph.getParagraphFormat().setBeforeSpacing(6);
		paragraph.appendText("Level 0");
		//Apply the default numbered list formats. 
		paragraph.getListFormat().applyDefNumberedStyle();
		paragraph=section.addParagraph();
		paragraph.appendText("Level 1");
		//Specify the list format to continue from the last list.
		paragraph.getListFormat().continueListNumbering();
		//Increment the list level.
		paragraph.getListFormat().increaseIndentLevel();
		paragraph=section.addParagraph();
		paragraph.appendText("Level 0");
		//Decrement the list level.
		paragraph.getListFormat().continueListNumbering();
		paragraph.getListFormat().decreaseIndentLevel();
		//Write the default bulleted list. 
		section.addParagraph();
		paragraph=section.addParagraph();	
		paragraph.appendText("Level 0");
		//Apply the default bulleted list formats.
		paragraph.getListFormat().applyDefBulletStyle();
		paragraph=section.addParagraph();
		paragraph.appendText("Level 1");
		//Specify the list format to continue from the last list.
		paragraph.getListFormat().continueListNumbering();
		//Increment the list level.
		paragraph.getListFormat().increaseIndentLevel();
		paragraph=section.addParagraph();
		paragraph.appendText("Level 0");
		//Specify the list format to continue from the last list.
		paragraph.getListFormat().continueListNumbering();
		//Decrement the list level.
		paragraph.getListFormat().decreaseIndentLevel();
		section.addParagraph();
		//Save the document in the given name and format.
		document.save("Sample.docx",FormatType.Docx);
		//Release the resources occupied by the WordDocument instance.
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
