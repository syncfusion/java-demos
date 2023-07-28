package styles;

import com.syncfusion.docio.*;
import com.syncfusion.javahelper.system.*;
import com.syncfusion.javahelper.system.drawing.*;

public class Styles {
	public static void main(String[] args) throws Exception {
		// Creates a new document.
		WordDocument document = new WordDocument();
		IWParagraphStyle style = null;
		// Adds a new section to the document
		WSection section = (WSection)document.addSection();
		// Sets margin of the section 
		section.getPageSetup().getMargins().setAll((float) 72);
		// Adds paragraph to the section.
		IWParagraph par = document.getLastSection().addParagraph();
		// Append text to the paragraph.
		WTextRange range = (WTextRange)par.appendText("Using CustomStyles");
		// Applies character format to text.
		WCharacterFormat characterFormat = range.getCharacterFormat();
		characterFormat.setTextBackgroundColor((ColorSupport.getBlue()).clone());
		characterFormat.setFontSize(18f);
		document.getLastParagraph().getParagraphFormat().setHorizontalAlignment(HorizontalAlignment.Center);
		// Creates Paragraph styles
		style = document.addParagraphStyle("MyStyle_Normal");
		characterFormat = style.getCharacterFormat();
		characterFormat.setFontName("Bitstream Vera Serif");
		characterFormat.setFontSize(10f);
		style.getParagraphFormat().setHorizontalAlignment(HorizontalAlignment.Justify);
		characterFormat.setTextColor((ColorSupport.fromArgb(0, 21, 84)).clone());

		style = document.addParagraphStyle("MyStyle_Low");
		characterFormat = style.getCharacterFormat();
		characterFormat.setFontName("Times New Roman");
		characterFormat.setFontSize(16f);
		characterFormat.setBold(true);
		
		style = document.addParagraphStyle("MyStyle_Medium");
		characterFormat = style.getCharacterFormat();
		characterFormat.setFontName("Monotype Corsiva");
		characterFormat.setFontSize(18f);
		characterFormat.setBold(true);
		characterFormat.setTextColor((ColorSupport.fromArgb(51, 66, 125)).clone());
		
		style = document.addParagraphStyle("Mystyle_High");
		characterFormat = style.getCharacterFormat();
		characterFormat.setFontName("Bitstream Vera Serif");
		characterFormat.setFontSize(20f);
		characterFormat.setBold(true);
		characterFormat.setTextColor((ColorSupport.fromArgb(242, 151, 50)).clone());
		IWParagraph paragraph = null;
		for (int i = 0; i < document.getStyles().getCount(); i++) {
			 // Skips to apply the document default styles and also paragraph style. 
			if ((document.getStyles().get(i).getName() == "Normal")
					|| (document.getStyles().get(i).getName() == "Default Paragraph Font")
					|| document.getStyles().get(i).getStyleType().getEnumValue() != StyleType.ParagraphStyle
							.getEnumValue())
				continue;
			// Gets styles from Document.
			style = (IWParagraphStyle) document.getStyles().get(i);
			// Adds a new paragraph
			section.addParagraph();
			paragraph = section.addParagraph();
			// Applies styles to the current paragraph.
			paragraph.applyStyle(style.getName());
			// Writes text with the current style and formatting.
			paragraph.appendText("Syncfusion Inc [" + style.getName() + "] Style");
			// Adds a new paragraph.
			section.addParagraph();
			paragraph = section.addParagraph();
			// Applies another style to the current paragraph.
			paragraph.applyStyle("MyStyle_Normal");
			// Writes text with current style.
			paragraph.appendText(
					"We are a leading provider of software components and tools for the Microsoft .NET platform. Powerful and feature-rich, Syncfusion's broad range of products are vital to mission-critical applications in organizations ranging from small businesses to multinational Fortune 100 companies. ");
		}
		// Save and close the document.
		document.save("Styles.docx");
		document.close();
		System.out.println("Word document generated successfully.");
	}

}
