package editequation;

import java.io.*;
import com.syncfusion.docio.*;
import com.syncfusion.javahelper.system.*;

public class EditEquation {
	public static void main(String[] args) throws Exception {
		// Initialize Word document.
		WordDocument document = new WordDocument();
		// Opens an existing template document..
		document.open(getDataDir("Mathematical Equation.docx"), FormatType.Docx);
		// Gets the last section in the document.
		WSection section = document.getLastSection();
		// Gets paragraph from Word document.
		WParagraph paragraph = (WParagraph) ObjectSupport
				.instanceOf(document.getLastSection().getBody().getChildEntities().get(3), WParagraph.class);
		// Gets MathML from the paragraph.
		WMath math = (WMath) ObjectSupport.instanceOf(paragraph.getChildEntities().get(0), WMath.class);
		// Gets the radical equation.
		IOfficeMathRadical mathRadical = (IOfficeMathRadical) ObjectSupport
				.instanceOf(math.getMathParagraph().getMaths().get(0).getFunctions().get(1), IOfficeMathRadical.class);
		// Gets the fraction equation in radical.
		IOfficeMathFraction mathFraction = (IOfficeMathFraction) ObjectSupport
				.instanceOf(mathRadical.getEquation().getFunctions().get(0), IOfficeMathFraction.class);
		// Gets the n-array in fraction.
		IOfficeMathNArray mathNAry = (IOfficeMathNArray) ObjectSupport
				.instanceOf(mathFraction.getNumerator().getFunctions().get(0), IOfficeMathNArray.class);
		// Gets the math script in n-array.
		IOfficeMathScript mathScript = (IOfficeMathScript) ObjectSupport
				.instanceOf(mathNAry.getEquation().getFunctions().get(0), IOfficeMathScript.class);
		// Gets the delimiter in math script.
		IOfficeMathDelimiter mathDelimiter = (IOfficeMathDelimiter) ObjectSupport
				.instanceOf(mathScript.getEquation().getFunctions().get(0), IOfficeMathDelimiter.class);
		// Gets the math script in delimiter.
		mathScript = (IOfficeMathScript) ObjectSupport
				.instanceOf(mathDelimiter.getEquation().get(0).getFunctions().get(0), IOfficeMathScript.class);
		// Gets the math run element in math script.
		IOfficeMathRunElement mathParaItem = (IOfficeMathRunElement) ObjectSupport
				.instanceOf(mathScript.getEquation().getFunctions().get(0), IOfficeMathRunElement.class);
		// Modifies the math text value.
		((WTextRange) ObjectSupport.instanceOf(mathParaItem.getItem(), WTextRange.class)).setText("x");
		// Gets the math bar in delimiter.
		IOfficeMathBar mathBar = (IOfficeMathBar) ObjectSupport
				.instanceOf(mathDelimiter.getEquation().get(0).getFunctions().get(2), IOfficeMathBar.class);
		// Gets the math run element in bar.
		mathParaItem = (IOfficeMathRunElement) ObjectSupport.instanceOf(mathBar.getEquation().getFunctions().get(0),
				IOfficeMathRunElement.class);
		// Modifies the math text value.
		((WTextRange) ObjectSupport.instanceOf(mathParaItem.getItem(), WTextRange.class)).setText("x");
		// Gets the paragraph from Word document.
		paragraph = (WParagraph) ObjectSupport.instanceOf(document.getLastSection().getBody().getChildEntities().get(6),
				WParagraph.class);
		// Gets MathML from the paragraph.
		math = (WMath) ObjectSupport.instanceOf(paragraph.getChildEntities().get(0), WMath.class);
		// Gets the math script equation.
		mathScript = (IOfficeMathScript) ObjectSupport
				.instanceOf(math.getMathParagraph().getMaths().get(0).getFunctions().get(0), IOfficeMathScript.class);
		// Gets the math run element in math script.
		mathParaItem = (IOfficeMathRunElement) ObjectSupport.instanceOf(mathScript.getEquation().getFunctions().get(0),
				IOfficeMathRunElement.class);
		// Modifies the math text value.
		((WTextRange) ObjectSupport.instanceOf(mathParaItem.getItem(), WTextRange.class)).setText("x");
		// Gets the paragraph from Word document.
		paragraph = (WParagraph) ObjectSupport.instanceOf(document.getLastSection().getBody().getChildEntities().get(7),
				WParagraph.class);
		// Gets MathML from the paragraph.
		WMath math2 = (WMath) ObjectSupport.instanceOf(paragraph.getChildEntities().get(0), WMath.class);
		// Gets bar equation.
		mathBar = (IOfficeMathBar) ObjectSupport
				.instanceOf(math2.getMathParagraph().getMaths().get(0).getFunctions().get(0), IOfficeMathBar.class);
		// Gets the math run element in bar.
		mathParaItem = (IOfficeMathRunElement) ObjectSupport.instanceOf(mathBar.getEquation().getFunctions().get(0),
				IOfficeMathRunElement.class);
		// Gets the math text.
		((WTextRange) ObjectSupport.instanceOf(mathParaItem.getItem(), WTextRange.class)).setText("x");
		// Saves and closes the document.
		document.save("Edit Equation.docx");
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
