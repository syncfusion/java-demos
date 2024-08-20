package createequationusinglatex;

import java.io.File;
import com.syncfusion.docio.*;

public class CreateEquationUsingLateX {

	public static void main(String[] args) throws Exception {
		//Load the document.
		WordDocument document = new WordDocument(getDataDir("Create Equation.docx"), FormatType.Docx);		
		// Get the last section of the document.
		WSection section = document.getLastSection();
	    // Set all margins of the page to 72 points.
		section.getPageSetup().getMargins().setAll((float)72);
		// Add a new paragraph to the section.
		IWParagraph paragraph = section.addParagraph();
		// Add text to the paragraph and set its font size to 28 points.
		IWTextRange textRange = paragraph.appendText("Mathematical equation");
		textRange.getCharacterFormat().setFontSize((float)28);
		 // Center align the paragraph.
		paragraph.getParagraphFormat().setHorizontalAlignment(HorizontalAlignment.Center);
	    // Set the spacing after the paragraph to 12 points.
		paragraph.getParagraphFormat().setAfterSpacing((float)12);
		paragraph=section.addParagraph();
		// Append equation to the paragraph.
		paragraph.appendMath("f\\left(x\\right)={a}_{0}+\\sum_{n=1}^{\\infty}{\\left({a}_{n}\\cos{\\frac{n\\pi{x}}{L}}+{b}_{n}\\sin{\\frac{n\\pi{x}}{L}}\\right)}");
		//Save the word document.
		document.save("Sample.docx");
		//Close the word document.
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
		if (!(dir.toString().endsWith("samples")))
			dir = dir.getParentFile();
		dir = new File(dir, "resources");
		dir = new File(dir, path);
		if (dir.isDirectory() == false)
			dir.mkdir();
		return dir.toString();
	}

}
