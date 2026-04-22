package formfillingandprotection;

import java.io.*;
import java.text.SimpleDateFormat;
import java.time.LocalDateTime;
import com.syncfusion.docio.*;
import com.syncfusion.javahelper.system.*;;

public class FormFillingandProtection {
	public static void main(String[] args) throws Exception {
		// Creates an Word document instance.
		WordDocument document = new WordDocument();
		// Loads the template document.
		document.open(getDataDir("ContentControlTemplate.docx"), FormatType.Docx);
		IWTextRange textRange;
		// Gets table from the template document.
		IWTable table = document.getLastSection().getTables().get(0);
		WTableRow row = table.getRows().get(1);
		IWParagraph cellPara = row.getCells().get(0).getParagraphs().get(0);
		// Calendar content control.
		// Accesses the date picker content control.
		IInlineContentControl inlineControl = (IInlineContentControl)cellPara.getChildEntities().get(2);
		textRange = (WTextRange)inlineControl.getParagraphItems().get(0);
		// Sets today's date to display.
		textRange.setText(DateTimeSupport.toString(LocalDateTime.now(), "d"));
		textRange.getCharacterFormat().setFontSize((float) 14);
		// Protects the content control.
		inlineControl.getContentControlProperties().setLockContents(true);
		// Plain text content controls
		table = document.getLastSection().getTables().get(1);
		row = table.getRows().get(0);
		cellPara = row.getCells().get(0).getLastParagraph();
		// Accesses the plain text content control.
		inlineControl = (IInlineContentControl)cellPara.getChildEntities().get(1);
		// Applies text formatting.
		applyTextFormatting(inlineControl, textRange, "Northwind Analytics");
		cellPara = row.getCells().get(1).getLastParagraph();
		inlineControl = (IInlineContentControl)cellPara.getChildEntities().get(1);
		// Applies text formatting.
		applyTextFormatting(inlineControl, textRange, "Northwind");
		row = table.getRows().get(1);
		cellPara = row.getCells().get(0).getLastParagraph();
		inlineControl = (IInlineContentControl)cellPara.getChildEntities().get(1);
		// Applies text formatting.
		applyTextFormatting(inlineControl, textRange, "10");
		cellPara = row.getCells().get(1).getLastParagraph();
		inlineControl = (IInlineContentControl)cellPara.getChildEntities().get(1);
		// Applies text formatting.
		applyTextFormatting(inlineControl, textRange, "Nancy Davolio");
		// CheckBox Content control
		row = table.getRows().get(2);
		cellPara = row.getCells().get(0).getLastParagraph();
		// Inserts checkbox content control.
		inlineControl = cellPara.appendInlineContentControl(ContentControlType.CheckBox);
		// Applies ContentControl formatting.
		applyContentControlFormatting(inlineControl);
		textRange = cellPara.appendText("C#, ");
		textRange.getCharacterFormat().setFontSize((float) 14);
		inlineControl = cellPara.appendInlineContentControl(ContentControlType.CheckBox);
		// Applies ContentControl formatting.
		applyContentControlFormatting(inlineControl);
		textRange = cellPara.appendText("VB");
		textRange.getCharacterFormat().setFontSize((float) 14);
		// Drop down list content control.
		cellPara = row.getCells().get(1).getLastParagraph();
		inlineControl = (IInlineContentControl)cellPara.getChildEntities().get(1);
		// Protects the inline content control
		inlineControl.getContentControlProperties().setLockContents(true);
		textRange = (WTextRange)inlineControl.getParagraphItems().get(0);
		textRange.setText("ASP.NET");
		textRange.getCharacterFormat().setFontSize((float) 14);
		inlineControl.getParagraphItems().add(textRange);
		ContentControlListItem item =listItem(inlineControl);
		// Adds items to the dropdown list.
		
		inlineControl.getContentControlProperties().getContentControlListItems().add(item);
		row = table.getRows().get(3);
		cellPara = row.getCells().get(0).getLastParagraph();
		// Calendar content control
		inlineControl = (IInlineContentControl)cellPara.getChildEntities().get(1);
		// Protects the inline content control
		inlineControl.getContentControlProperties().setLockContents(true);
		textRange = (WTextRange)inlineControl.getParagraphItems().get(0);
		textRange.setText(DateTimeSupport.toString(DateTimeSupport.addDays(LocalDateTime.now(), (double) -5), "d"));
		textRange.getCharacterFormat().setFontSize((float) 14);
		cellPara = row.getCells().get(1).getLastParagraph();
		inlineControl = (IInlineContentControl)cellPara.getChildEntities().get(1);
		// Protects the inline content control
		inlineControl.getContentControlProperties().setLockContents(true);
		textRange = (WTextRange)inlineControl.getParagraphItems().get(0);
		textRange.setText(
				DateTimeSupport.toString(DateTimeSupport.addDays(LocalDateTime.now(), (double) ((double) 10)), "d"));
		textRange.getCharacterFormat().setFontSize((float) 14);
		// Accesses the block content control.
		BlockContentControl blockContentControl = (BlockContentControl)((WSection)document.getChildEntities().get(0))
						.getBody().getChildEntities().get(2);
		// Protects the block content control
		blockContentControl.getContentControlProperties().setLockContents(true);
		// Save and close the document.
		document.save("Form Filling and Protection.docx");
		document.close();
		System.out.println("Word document generated successfully.");
	}

	/**
	 * Apply text formatting.
	 * 
	 * @param inlineControl Represents a inline content control.
	 * @param textRange     Represents the text to be formatted
	 * @param text          textRange Represents the text to be appended
	 * 
	 */
	static IInlineContentControl applyTextFormatting(IInlineContentControl inlineControl, IWTextRange textRange,
			String text) throws Exception {
		// Protects the content control.
		inlineControl.getContentControlProperties().setLockContents(true);
		textRange = (WTextRange)inlineControl.getParagraphItems().get(0);
		// Sets text in plain text content control.
		textRange.setText(text);
		textRange.getCharacterFormat().setFontSize((float) 14);
		return inlineControl;
	}

	/**
	 * Apply ContentControl formatting.
	 * 
	 * @param inlineControl Represents a inline content control to be formatted.
	 *
	 */
	static IInlineContentControl applyContentControlFormatting(IInlineContentControl inlineControl) throws Exception {
		inlineControl.getContentControlProperties().setLockContents(true);
		// Sets checkbox as checked state.
		inlineControl.getContentControlProperties().setIsChecked(true);
		return inlineControl;
	}

	/**
	 * Adds items to the dropdown list.
	 * 
	 * @param inlineControl inlineControl Represents a inline content control to be formatted.	 * 
	 */
	static ContentControlListItem listItem( IInlineContentControl inlineControl)
			throws Exception {
		// Creates an item for dropdown list.
		ContentControlListItem item;
		item = new ContentControlListItem();
		// Sets the text to be displayed as list item.
		item.setDisplayText("ASP.NET MVC");
		item.setValue("2");
		// Adds item to the dropdown list.
		inlineControl.getContentControlProperties().getContentControlListItems().add(item);
		item = new ContentControlListItem();
		// Sets the text to be displayed as list item.
		item.setDisplayText("Windows Forms");
		item.setValue("3");
		// Adds item to the dropdown list.
		inlineControl.getContentControlProperties().getContentControlListItems().add(item);
		item = new ContentControlListItem();
		// Sets the text to be displayed as list item.
		item.setDisplayText("WPF");
		item.setValue("4");
		// Adds item to the dropdown list.
		inlineControl.getContentControlProperties().getContentControlListItems().add(item);
		item = new ContentControlListItem();
		// Sets the text to be displayed as list item.
		item.setDisplayText("Xamarin");
		item.setValue("5");
		return item;
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
