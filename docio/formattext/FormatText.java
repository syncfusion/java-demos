package formattext;

import java.lang.reflect.Array;
import com.syncfusion.docio.*;
import com.syncfusion.javahelper.system.*;
import com.syncfusion.javahelper.system.drawing.*;

public class FormatText {
	static String[] products = new String[] { "Tools", "Grid", "Chart", "Edit", "Diagram", "XlsIO", "Grouping",
			"Calculate", "PDF", "HTMLUI", "DocIO" };
	static String[] forms = new String[] { "Windows", "Web" };
	static IWSection section1;
	static IWParagraph paragraph = null;
	static IWTextRange textRange = null;

	public static void main(String[] args) throws Exception {
		// List of FontNames.
		String[] fontNames = { "Arial", "Times New Roman", "Monotype Corsiva", " Book Antiqua ", "Bitstream Vera Sans",
				"Comic Sans MS", "Microsoft Sans Serif", "Batang" };
		// Creates a new document.
		WordDocument document = new WordDocument();
		// Adds a new section to the document.
		IWSection section = document.addSection();
		section.getPageSetup().getMargins().setAll((float) 72);
		// Adds a new paragraph to the section.
		IWParagraph paragraph = section.addParagraph();
		// Append text to the paragraph.
		paragraph.appendText(
				"This sample demonstrates various text and paragraph formatting supported by Essential DocIO.");
		section.addParagraph();
		section.addParagraph();
		section = document.addSection();
		// Sets break type to the section.
		section.setBreakCode(SectionBreakCode.NoBreak);
		// Adds two columns to the section.
		section.addColumn((float) 250, (float) 20);
		section.addColumn((float) 250, (float) 20);
		// Text Formatting.
		IWTextRange text = null;
		WCharacterFormat characterFormat;
		// Writing Text with different Formatting styles.
		for (int i = 8, j = 0, k = 0; i <= 20; i++, j++, k++) {
			if (j >= Array.getLength(fontNames))
				j = 0;
			paragraph = section.addParagraph();
			text = paragraph.appendText("Essential DocIO " + "[" + fontNames[(int) j] + "]");
			// Applies character format to text.
			characterFormat = text.getCharacterFormat();
			characterFormat.setFontName(fontNames[(int) j]);
			characterFormat.setUnderlineStyle(UnderlineStyle.valueOf(k));
			characterFormat.setFontSize((float) i);
			characterFormat.setTextColor((ColorSupport.fromArgb(128, 128, 128)).clone());
		}
		// Adds paragraph to the section.
		section.addParagraph();
		paragraph.getParagraphFormat().setColumnBreakAfter(true);
		paragraph = section.addParagraph();
		text = paragraph.appendText("More formatting Options List...");
		// Applies character format to text.
		characterFormat = text.getCharacterFormat();
		characterFormat.setFontSize((float) 18);
		characterFormat.setFontName(fontNames[(int) 2]);
		section.addParagraph();
		paragraph = section.addParagraph();
		// Applies character formatting to the textrange.
		applyCharacterFormatting(paragraph);
		section = document.addSection();
		// Sets break type to the section.
		section.setBreakCode(SectionBreakCode.NoBreak);
		paragraph = section.addParagraph();
		paragraph.appendText("Following paragraphs illustrates various paragraph formattings");
		// Applies paragraph formattings.
		applyParagraphFormatting(paragraph, section);
		paragraph.appendText("KEEP TOGETHER END").getCharacterFormat().setBold(true);
		section = document.addSection();
		// Sets margins to section.
		MarginsF margins = section.getPageSetup().getMargins();
		margins.setTop((float) 20);
		margins.setBottom((float) 20);
		margins.setLeft((float) 50);
		margins.setRight((float) 20);
		// Adds paragraph to section.
		paragraph = section.addParagraph();
		paragraph.appendText(
				"This document demonstrates the Bullets and Numbering functionality available in Essential DocIO");
		section1 = document.addSection();
		section1.getColumns().add(new Column(document));
		section1.getColumns().add(new Column(document));
		section1.makeColumnsEqual();
		section1.setBreakCode(SectionBreakCode.NoBreak);
		productDetailsInBullets();
		productDetailsInNumbers();
		// Save and close the document.
		document.save("Format Text.docx");
		document.close();
		System.out.println("Word document generated successfully.");
	}

	/**
	 * 
	 * Product details in bullets.
	 */
	public static void productDetailsInBullets() throws Exception {
		// Adds paragraph to section.
		section1.addParagraph();
		paragraph = section1.addParagraph();
		// Append text to the paragraph.
		textRange = paragraph.appendText("List of Syncfusion products.");
		paragraph.getListFormat().applyDefBulletStyle();
		WCharacterFormat characterFormat = textRange.getCharacterFormat();
		// Applies character format to text.
		characterFormat.setFontSize((float) 15);
		characterFormat.setFontName("Monotype Corsiva");
		section1.addParagraph();
		// Sets bullet list to paragraphs.
		for (Object s_tempObj : products) {
			String s = (String) s_tempObj;
			section1.addParagraph();
			paragraph = section1.addParagraph();
			// Append text to the paragraph.
			paragraph.appendText(String.valueOf("Essential ").concat(String.valueOf(s)));
			paragraph.getListFormat().continueListNumbering();
			paragraph.getListFormat().setListLevelNumber(1);
			section1.addParagraph();
			for (Object s1_tempObj : forms) {
				String s1 = (String) s1_tempObj;
				if (s.equals("HTMLUI")) {
					paragraph = section1.addParagraph();
					// Append text to the paragraph.
					paragraph.appendText(String.valueOf(s1).concat(String.valueOf(" Forms")));
					paragraph.getListFormat().continueListNumbering();
					paragraph.getListFormat().setListLevelNumber(2);
					break;
				} else

				{
					paragraph = section1.addParagraph();
					// Append text to the paragraph.
					paragraph.appendText(String.valueOf(s1).concat(String.valueOf(" Forms")));
					paragraph.getListFormat().continueListNumbering();
					paragraph.getListFormat().setListLevelNumber(2);
				}
			}
		}
	}

	/**
	 * 
	 * Product details in numbers.
	 */
	public static void productDetailsInNumbers() throws Exception {
		// Adds paragraph to section.
		section1.addParagraph();
		paragraph = section1.addParagraph();
		// Append text to the paragraph.
		textRange = paragraph.appendText("List of Syncfusion products.");
		paragraph.getListFormat().applyDefNumberedStyle();
		// Applies character format to text.
		WCharacterFormat characterFormat = textRange.getCharacterFormat();
		characterFormat.setFontSize((float) 15);
		characterFormat.setFontName("Monotype Corsiva");
		// Adds paragraph to section.
		section1.addParagraph();
		// Sets number list to paragraphs.
		for (Object s_tempObj : products) {
			String s = (String) s_tempObj;
			section1.addParagraph();
			paragraph = section1.addParagraph();
			// Append text to the paragraph.
			paragraph.appendText(String.valueOf("Essential ").concat(String.valueOf(s)));
			paragraph.getListFormat().continueListNumbering();
			paragraph.getListFormat().setListLevelNumber(1);
			section1.addParagraph();
			for (Object s1_tempObj : forms) {
				String s1 = (String) s1_tempObj;
				if (s.equals("HTMLUI")) {
					paragraph = section1.addParagraph();
					// Append text to the paragraph.
					paragraph.appendText(String.valueOf(s1).concat(String.valueOf(" Forms")));
					paragraph.getListFormat().continueListNumbering();
					paragraph.getListFormat().setListLevelNumber(2);
					break;
				} else

				{
					paragraph = section1.addParagraph();
					// Append text to the paragraph.
					paragraph.appendText(String.valueOf(s1).concat(String.valueOf(" Forms")));
					paragraph.getListFormat().continueListNumbering();
					paragraph.getListFormat().setListLevelNumber(2);
				}
			}
		}
	}

	/**
	 * Apply character formatting to the textrange.
	 * 
	 * @param paragraph Represents the paragraph in which the formatting is applied.
	 * 
	 */
	static IWParagraph applyCharacterFormatting(IWParagraph paragraph) throws Exception {
		paragraph.appendText("AllCaps \n\n").getCharacterFormat().setAllCaps(true);
		paragraph.appendText("Bold \n\n").getCharacterFormat().setBold(true);
		paragraph.appendText("DoubleStrike \n\n").getCharacterFormat().setDoubleStrike(true);
		paragraph.appendText("Emboss \n\n").getCharacterFormat().setEmboss(true);
		paragraph.appendText("Engrave \n\n").getCharacterFormat().setEngrave(true);
		paragraph.appendText("Italic \n\n").getCharacterFormat().setItalic(true);
		paragraph.appendText("Shadow \n\n").getCharacterFormat().setShadow(true);
		paragraph.appendText("SmallCaps \n\n").getCharacterFormat().setSmallCaps(true);
		paragraph.appendText("Strikeout \n\n").getCharacterFormat().setStrikeout(true);
		paragraph.appendText("Some Text");
		paragraph.appendText("SubScript \n\n").getCharacterFormat().setSubSuperScript(SubSuperScript.SubScript);
		paragraph.appendText("Some Text");
		paragraph.appendText("SuperScript \n\n").getCharacterFormat().setSubSuperScript(SubSuperScript.SuperScript);
		paragraph.appendText("TextBackgroundColor \n\n").getCharacterFormat()
				.setTextBackgroundColor((ColorSupport.getLightBlue()).clone());

		return paragraph;
	}

	/**
	 * Apply paragraph formatting to the textrange.
	 * 
	 * @param section   Represents a section.
	 * @param paragraph Represents the paragraph in which the formatting is applied
	 */
	static IWParagraph applyParagraphFormatting(IWParagraph paragraph, IWSection section) throws Exception {
		// Adds paragraph to the section.
		paragraph = section.addParagraph();
		paragraph.appendText(
				"We will use this paragraph to illustrate several Microsoft Word features using Essential DocIO. It will be used to illustrate Space Before, Space After, and Line Spacing. Space Before tells Microsoft Word how much space to leave before the paragraph. Space After tells Microsoft Word how much space to leave after the paragraph. Line Spacing sets the space between lines within a paragraph.It also explains about first line indentation,backcolor and paragraph border.");
		// Applies paragraph format. 
		WParagraphFormat paragraphFormat = paragraph.getParagraphFormat();
		paragraphFormat.setBeforeSpacing(20f);
		paragraphFormat.setAfterSpacing(30f);
		paragraphFormat.setBackColor((ColorSupport.getLightGray()).clone());
		paragraphFormat.getBorders().setBorderType(BorderStyle.Single);
		paragraphFormat.setFirstLineIndent(20f);
		paragraphFormat.setLineSpacing(20f);
		// Adds paragraph to the section.
		paragraph = section.addParagraph();
		paragraph.appendText(
				"This is a sample paragraph. It is used to illustrate alignment. Left-justified text is aligned on the left. Right-justified text is aligned with on the right. Centered text is centered between the left and right margins. You can use Center to center your titles. Justified text is flush on both sides.");
		// Applies paragraph format.
		paragraphFormat = paragraph.getParagraphFormat();
		paragraphFormat.setLineSpacing(20f);
		paragraphFormat.setHorizontalAlignment(HorizontalAlignment.Right);
		paragraphFormat.setLineSpacingRule(LineSpacingRule.Exactly);
		// Adds paragraph to the section.
		section.addParagraph();
		paragraph = section.addParagraph();
		paragraph.getParagraphFormat().setKeep(true);
		paragraph.appendText("KEEP TOGETHER").getCharacterFormat().setBold(true);
		paragraph.appendText(
				"This is a sample paragraph. It is used to illustrate Keep together of MsWord. You can control where Microsoft Word positions automatic page breaks (page break: The point at which one page ends and another begins. Microsoft Word inserts an 'automatic' (or soft) page break for you, or you can force a page break at a specific location by inserting a 'manual' (or hard) page break.) by setting pagination options.It keeps the lines in a paragraph together when there is page break")
				.getCharacterFormat().setFontSize(12f);
		for (int i = 0; i < 10; i++) {
			paragraph.appendText("Text Body_Text Body_Text Body_Text Body_Text Body_Text Body_Text Body")
					.getCharacterFormat().setFontSize(12f);
			paragraph.getParagraphFormat().setLineSpacing(20f);

		}
		return paragraph;
	}
}
