package createequation;

import java.io.*;
import com.syncfusion.docio.*;
import com.syncfusion.javahelper.system.*;

public class CreateEquation {
	public static void main(String[] args) throws Exception {
		// Creates a new word document.
		WordDocument document = new WordDocument();
		// Adds new section to the document.
		IWSection section = document.addSection();
		// Sets page margins.
		document.getLastSection().getPageSetup().getMargins().setAll((float) 72);
		// Adds new paragraph to the section.
		IWParagraph paragraph = section.addParagraph();
		// Appends text to paragraph.
		IWTextRange textRange = paragraph.appendText("Mathematical equations");
		textRange.getCharacterFormat().setFontSize((float) 28);
		paragraph.getParagraphFormat().setHorizontalAlignment(HorizontalAlignment.Center);
		paragraph.getParagraphFormat().setAfterSpacing((float) 12);
		// Adds new paragraph to the section.
		paragraph = addParagraph(section, "This is an expansion of the sum (1+X) to the power of n.");
		// Creates an equation with sum to the power of N.
		createSumToThePowerOfN(paragraph);
		// Adds new paragraph to the section.
		paragraph = addParagraph(section, "This is a Fourier series for the function of period 2L");
		// Creates a Fourier series equation.
		createFourierseries(paragraph);
		// Adds new paragraph to the section.
		paragraph = addParagraph(section, "This is an expansion of triple scalar product");
		// Creates a triple scalar product equation.
		createTripleScalarProduct(paragraph);
		// Adds new paragraph to the section.
		paragraph = addParagraph(section, "This is an expansion of gamma function");
		// Creates a gamma function equation.
		createGammaFunction(paragraph);
		// Adds new paragraph to the section.
		paragraph = addParagraph(section, "This is an expansion of vector relation ");
		// Creates a vector relation equation.
		createVectorRelation(paragraph);
		// Saves and closes the document.
		document.save("Create Equation.docx");
		document.close();
		System.out.println("Word document generated successfully.");
	}

	/**
	 * Adds new paragraph into the section.
	 * 
	 * @param section Represents a section in Word document
	 * @param text    Represents a text to append in paragraph
	 * @return Returns a paragraph to add equation
	 */
	private static IWParagraph addParagraph(IWSection section, String text) throws Exception {
		// Adds new paragraph to the section.
		IWParagraph paragraph = section.addParagraph();
		// Adds new paragraph to add text.
		paragraph = section.addParagraph();
		// Appends text to paragraph.
		paragraph.appendText(text);
		paragraph.getParagraphFormat().setAfterSpacing((float) 12);
		paragraph.getParagraphFormat().setBeforeSpacing((float) 12);
		// Adds new paragraph to add equation.
		paragraph = section.addParagraph();
		paragraph.getParagraphFormat().setAfterSpacing((float) 12);
		return paragraph;
	}

	/**
	 * Creates an expansion of triple scalar product
	 * 
	 * @param paragraph Represents a paragraph to add MathML element
	 */
	private static void createTripleScalarProduct(IWParagraph paragraph) throws Exception {
		WordDocument document = paragraph.getDocument();
		// Creates a MathML element.
		WMath math = paragraph.appendMath();
		// Adds an office math.
		IOfficeMath officeMath = math.getMathParagraph().getMaths().add();
		// Unicode value of middle dot.
		String middleDot = "\u00B7";
		String multiplicationSign = "\u00D7";
		String text = StringSupport.concat(StringSupport
				.concat(StringSupport.concat(StringSupport.concat("A", middleDot), "B"), multiplicationSign), "C");
		// Adds a math item.
		IOfficeMathRunElement officeMathParaItem = addMathText(document, officeMath, text);
		// Sets style for math text.
		officeMathParaItem.getMathFormat().setStyle(MathStyleType.Bold);
		// Adds math text.
		officeMathParaItem = addMathText(document, officeMath, "=");
		// Sets style for math text.
		officeMathParaItem.getMathFormat().setStyle(MathStyleType.Bold);
		text = StringSupport.concat(StringSupport
				.concat(StringSupport.concat(StringSupport.concat("A", multiplicationSign), "B"), middleDot), "C");
		// Adds math text.
		officeMathParaItem = addMathText(document, officeMath, text);
		// Sets style for math text.
		officeMathParaItem.getMathFormat().setStyle(MathStyleType.Bold);
		// Adds math text.
		officeMathParaItem = addMathText(document, officeMath, "=");
		// Adds a delimiter.
		IOfficeMathDelimiter mathDelimiter = (IOfficeMathDelimiter) ObjectSupport
				.instanceOf(officeMath.getFunctions().add(MathFunctionType.Delimiter), IOfficeMathDelimiter.class);
		// Sets begin character for delimiter.
		mathDelimiter.setBeginCharacter("|");
		// Sets end character for delimiter.
		mathDelimiter.setEndCharacter("|");
		// Apply formats for delimiter.
		mathDelimiter.setControlProperties(new WCharacterFormat(document));
		((WCharacterFormat) ObjectSupport.instanceOf(mathDelimiter.getControlProperties(), WCharacterFormat.class))
				.setItalic(true);
		// Adds arguments for delimiter.
		officeMath = (IOfficeMath) ObjectSupport.instanceOf(mathDelimiter.getEquation().add(), IOfficeMath.class);
		// Add matrix into delimiter.
		IOfficeMathMatrix mathMatrix = (IOfficeMathMatrix) ObjectSupport
				.instanceOf(officeMath.getFunctions().add(MathFunctionType.Matrix), IOfficeMathMatrix.class);
		mathMatrix.getColumns().add();
		mathMatrix.getColumns().add();
		mathMatrix.getColumns().add();
		// Adds a new row.
		IOfficeMathMatrixRow mathMatrixRow = mathMatrix.getRows().add();
		// Add values to row.
		addMatrixRowValues(document, mathMatrixRow, "A");
		// Adds a new row.
		mathMatrixRow = mathMatrix.getRows().add();
		// Add values to row.
		addMatrixRowValues(document, mathMatrixRow, "B");
		// Adds a new row.
		mathMatrixRow = mathMatrix.getRows().add();
		// Add values to row.
		addMatrixRowValues(document, mathMatrixRow, "C");
	}

	/**
	 * Creates an expansion of vector relation
	 * 
	 * @param paragraph Represents a paragraph to add MathML element
	 */
	private static void createVectorRelation(IWParagraph paragraph) throws Exception {
		WordDocument document = paragraph.getDocument();
		// Creates a MathML element.
		WMath math = paragraph.appendMath();
		IOfficeMath officeMath = math.getMathParagraph().getMaths().add();
		// Adds an accent equation.
		addMathAccent(document, officeMath, (short) 8407, "A");
		// Adds a math text.
		String middleDot = "\u00B7";
		officeMath = math.getMathParagraph().getMaths().add();
		IOfficeMathRunElement officeMathParaItem = addMathText(document, officeMath, middleDot);
		// Adds an accent equation.
		officeMath = math.getMathParagraph().getMaths().add();
		addMathAccent(document, officeMath, (short) 8407, "B");
		officeMath = math.getMathParagraph().getMaths().add();
		// Adds a math text.
		String multiplicationSign = "\u00D7";
		officeMathParaItem = addMathText(document, officeMath, multiplicationSign);
		// Adds an accent equation.
		officeMath = math.getMathParagraph().getMaths().add();
		addMathAccent(document, officeMath, (short) 8407, "C");
		// Adds a math text.
		officeMath = math.getMathParagraph().getMaths().add();
		officeMathParaItem = addMathText(document, officeMath, "=");
		// Adds an accent equation.
		officeMath = math.getMathParagraph().getMaths().add();
		addMathAccent(document, officeMath, (short) 8407, "A");
		// Adds a math text.
		officeMath = math.getMathParagraph().getMaths().add();
		officeMathParaItem = addMathText(document, officeMath, multiplicationSign);
		// Adds an accent equation.
		officeMath = math.getMathParagraph().getMaths().add();
		addMathAccent(document, officeMath, (short) 8407, "B");
		// Adds a math text.
		officeMath = math.getMathParagraph().getMaths().add();
		officeMathParaItem = addMathText(document, officeMath, middleDot);
		// Adds an accent equation.
		officeMath = math.getMathParagraph().getMaths().add();
		addMathAccent(document, officeMath, (short) 8407, "C");
	}

	/**
	 * Converts short value to string
	 * 
	 * @param value Represents a short value
	 * @return Returns string value
	 */
	private static String convertShortToString(short value) throws Exception {
		char chrValue = ConvertSupport.toChar(value);
		String strValue = ConvertSupport.toString(chrValue);
		return strValue;
	}

	/**
	 * Creates an equation with sum to the power of N
	 * 
	 * @param paragraph Represents a paragraph to add MathML element
	 */
	private static void createSumToThePowerOfN(IWParagraph paragraph) throws Exception {
		WordDocument document = paragraph.getDocument();
		// Creates a new MathML element.
		WMath math = paragraph.appendMath();
		IOfficeMath officeMath = math.getMathParagraph().getMaths().add();
		// Adds a math script element.
		IOfficeMathScript mathScript = addMathScript(officeMath, MathScriptType.Superscript);
		// Adds a delimiter.
		IOfficeMathDelimiter mathDelimiter = (IOfficeMathDelimiter) ObjectSupport.instanceOf(
				mathScript.getEquation().getFunctions().add(MathFunctionType.Delimiter), IOfficeMathDelimiter.class);
		// Adds an office math in the delimiter.
		officeMath = (IOfficeMath) ObjectSupport.instanceOf(mathDelimiter.getEquation().add(), IOfficeMath.class);
		// Adds a math text.
		IOfficeMathRunElement mathParaItem = addMathText(document, officeMath, "1+x");
		mathParaItem = addMathText(document, mathScript.getScript(), "n");
		officeMath = math.getMathParagraph().getMaths().add(1);
		mathParaItem = addMathText(document, officeMath, "=1+");
		// Adds a math fraction.
		IOfficeMathFraction mathFraction = (IOfficeMathFraction) ObjectSupport.instanceOf(
				math.getMathParagraph().getMaths().add().getFunctions().add(MathFunctionType.Fraction),
				IOfficeMathFraction.class);
		// Adds a numerator text.
		addMathText(document, mathFraction.getNumerator(), "nx");
		// Adds a denominator text.
		addMathText(document, mathFraction.getDenominator(), "1!");
		// Adds a math text.
		officeMath = math.getMathParagraph().getMaths().add();
		// Adds a math fraction.
		mathParaItem = addMathText(document, officeMath, "+");
		mathFraction = (IOfficeMathFraction) ObjectSupport.instanceOf(
				math.getMathParagraph().getMaths().add().getFunctions().add(MathFunctionType.Fraction),
				IOfficeMathFraction.class);
		// Adds a numerator text.
		addMathText(document, mathFraction.getNumerator(), "n");
		// Adds a delimiter.
		mathDelimiter = (IOfficeMathDelimiter) ObjectSupport.instanceOf(
				mathFraction.getNumerator().getFunctions().add(MathFunctionType.Delimiter), IOfficeMathDelimiter.class);
		// Adds a math text for delimiter.
		officeMath = (IOfficeMath) ObjectSupport.instanceOf(mathDelimiter.getEquation().add(), IOfficeMath.class);
		addMathText(document, officeMath, "n-1");
		// Adds a math script.
		mathScript = (IOfficeMathScript) ObjectSupport.instanceOf(
				mathFraction.getNumerator().getFunctions().add(MathFunctionType.SubSuperscript),
				IOfficeMathScript.class);
		// Adds a math text for Superscript.
		addMathText(document, mathScript.getEquation(), "x");
		addMathText(document, mathScript.getScript(), "2");
		// Adds a denominator text
		addMathText(document, mathFraction.getDenominator(), "2!");
		// Adds a math text.
		officeMath = math.getMathParagraph().getMaths().add();
		addMathText(document, officeMath, "+...");
	}

	/**
	 * Creates an expansion of Gamma function
	 * 
	 * @param paragraph Represents a paragraph to add MathML element
	 */
	private static void createGammaFunction(IWParagraph paragraph) throws Exception {
		WordDocument document = paragraph.getDocument();
		// Creates a new MathML element.
		WMath math = paragraph.appendMath();
		IOfficeMath officeMath = math.getMathParagraph().getMaths().add();
		String capitalGamma = "\u0393";
		// Adds a math text.
		IOfficeMathRunElement officeMathParaItem = addMathText(document, officeMath, capitalGamma);
		officeMathParaItem.getMathFormat().setStyle(MathStyleType.Regular);
		IOfficeMathDelimiter mathDelimiter = (IOfficeMathDelimiter) ObjectSupport.instanceOf(
				math.getMathParagraph().getMaths().add().getFunctions().add(MathFunctionType.Delimiter),
				IOfficeMathDelimiter.class);
		officeMath = mathDelimiter.getEquation().add();
		// Adds a math text.
		officeMathParaItem = addMathText(document, officeMath, "z");
		officeMath = math.getMathParagraph().getMaths().add();
		// Adds a math text.
		officeMathParaItem = addMathText(document, officeMath, "=");
		// Adds an n array element.
		IOfficeMathNArray mathNAry = (IOfficeMathNArray) ObjectSupport.instanceOf(
				math.getMathParagraph().getMaths().add().getFunctions().add(MathFunctionType.NArray),
				IOfficeMathNArray.class);
		// Adds a math text.
		addMathText(document, mathNAry.getSubscript(), "0");
		String infinitySymbol = "\u221E";
		// Adds a math text.
		addMathText(document, mathNAry.getSuperscript(), infinitySymbol);
		// Adds a math superscript.
		IOfficeMathScript mathScript = addMathScript(mathNAry.getEquation(), MathScriptType.Superscript);
		// Adds a math text for Superscript.
		addMathText(document, mathScript.getEquation(), "t");
		addMathText(document, mathScript.getScript(), "z-1");
		// Adds a Superscript.
		mathScript = addMathScript(mathNAry.getEquation(), MathScriptType.Superscript);
		addMathText(document, mathScript.getEquation(), "e");
		addMathText(document, mathScript.getScript(), "-t");
		// Adds a math text in n Array equation.
		addMathText(document, mathNAry.getEquation(), "dt");
		// Adds a math text.
		addMathText(document, math.getMathParagraph().getMaths().add(), "=");
		// Adds a fraction equation.
		IOfficeMathFraction mathFraction = (IOfficeMathFraction) ObjectSupport.instanceOf(
				math.getMathParagraph().getMaths().add().getFunctions().add(MathFunctionType.Fraction),
				IOfficeMathFraction.class);
		// Adds a math script.
		mathScript = addMathScript(mathFraction.getNumerator(), MathScriptType.Superscript);
		// Adds a math text for Superscript.
		addMathText(document, mathScript.getEquation(), "e");
		addMathText(document, mathScript.getScript(), "-");
		// Unicode of small gamma.
		String smallGamma = "\u03B3";
		addMathText(document, mathScript.getScript(), smallGamma);
		addMathText(document, mathScript.getScript(), "z");
		// Adds a math text for denominator.
		addMathText(document, mathFraction.getDenominator(), "z");
		// Adds an n-array element.
		mathNAry = (IOfficeMathNArray) ObjectSupport.instanceOf(
				math.getMathParagraph().getMaths().add().getFunctions().add(MathFunctionType.NArray),
				IOfficeMathNArray.class);
		// Unicode value of n-array product.
		String symbol = "\u220F";
		// Adds a n-array character.
		mathNAry.setNArrayCharacter(symbol);
		// Adds a math text.
		addMathText(document, mathNAry.getSubscript(), "k=1");
		addMathText(document, mathNAry.getSuperscript(), infinitySymbol);
		// Adds a math script.
		mathScript = addMathScript(mathNAry.getEquation(), MathScriptType.Superscript);
		// Adds a math delimiter element.
		mathDelimiter = (IOfficeMathDelimiter) ObjectSupport.instanceOf(
				mathScript.getEquation().getFunctions().add(MathFunctionType.Delimiter), IOfficeMathDelimiter.class);
		// Adds a equation to the delimiter equation collection.
		officeMath = mathDelimiter.getEquation().add();
		// Adds a math text.
		addMathText(document, officeMath, "1+");
		// Adds a fraction element.
		mathFraction = (IOfficeMathFraction) ObjectSupport
				.instanceOf(officeMath.getFunctions().add(MathFunctionType.Fraction), IOfficeMathFraction.class);
		// Adds a math text for numerator.
		addMathText(document, mathFraction.getNumerator(), "z");
		// Adds a math text for denominator.
		addMathText(document, mathFraction.getDenominator(), "k");
		// Adds a math text.
		addMathText(document, mathScript.getScript(), "-1");
		// Adds a Superscript equation.
		mathScript = addMathScript(mathNAry.getEquation(), MathScriptType.Superscript);
		// Adds a math text for Superscript.
		addMathText(document, mathScript.getEquation(), "e");
		addMathText(document, mathScript.getScript(), "z");
		officeMathParaItem = addMathText(document, mathScript.getScript(), "/");
		officeMathParaItem.getMathFormat().setHasLiteral(true);
		addMathText(document, mathScript.getScript(), "k");
		// Adds a math text.
		addMathText(document, math.getMathParagraph().getMaths().add(), ",");
		officeMathParaItem = addMathText(document, math.getMathParagraph().getMaths().add(), "  ");
		// Sets style for math text.
		officeMathParaItem.getMathFormat().setStyle(MathStyleType.Regular);
		addMathText(document, math.getMathParagraph().getMaths().add(), smallGamma);
		String text = StringSupport.concat("\u2248", "0.577216");
		addMathText(document, math.getMathParagraph().getMaths().add(), text);
	}

	/**
	 * Creates a Fourier series equation
	 * 
	 * @param paragraph Represents a paragraph to add MathML element
	 */
	private static void createFourierseries(IWParagraph paragraph) throws Exception {
		// Creates a new MathML element.
		WordDocument document = paragraph.getDocument();
		WMath math = paragraph.appendMath();
		// Adds a math.
		IOfficeMath officeMath = math.getMathParagraph().getMaths().add();
		// Adds a math text.
		addMathText(document, officeMath, "f");
		// Adds a math delimiter.
		IOfficeMathDelimiter mathDelimiter = (IOfficeMathDelimiter) ObjectSupport.instanceOf(
				math.getMathParagraph().getMaths().add().getFunctions().add(MathFunctionType.Delimiter),
				IOfficeMathDelimiter.class);
		// Adds a math in the delimiter.
		officeMath = (IOfficeMath) ObjectSupport.instanceOf(mathDelimiter.getEquation().add(), IOfficeMath.class);
		// Adds a math text.
		addMathText(document, officeMath, "x");
		addMathText(document, math.getMathParagraph().getMaths().add(), "=");
		// Adds a Subscript equation.
		IOfficeMathScript mathScript = addMathScript(math.getMathParagraph().getMaths().add(),
				MathScriptType.Subscript);
		// Adds a math text.
		addMathText(document, mathScript.getEquation(), "a");
		addMathText(document, mathScript.getScript(), "0");
		addMathText(document, math.getMathParagraph().getMaths().add(), "+");
		// Adds a math n-array.
		IOfficeMathNArray mathNAry = (IOfficeMathNArray) ObjectSupport.instanceOf(
				math.getMathParagraph().getMaths().add().getFunctions().add(MathFunctionType.NArray),
				IOfficeMathNArray.class);
		// Unicode value of n-array summation.
		String sigma = "\u2211";
		// Sets the value as the n-array character.
		mathNAry.setNArrayCharacter(sigma);
		mathNAry.setHasGrow(true);
		// Adds a math text.
		addMathText(document, mathNAry.getSubscript(), "n=1");
		String infinitySymbol = "\u221E";
		// Adds a math text.
		addMathText(document, mathNAry.getSuperscript(), infinitySymbol);
		// Adds a math delimiter.
		mathDelimiter = (IOfficeMathDelimiter) ObjectSupport.instanceOf(
				mathNAry.getEquation().getFunctions().add(MathFunctionType.Delimiter), IOfficeMathDelimiter.class);
		// Adds an math in the delimiter equation collection.
		officeMath = (IOfficeMath) ObjectSupport.instanceOf(mathDelimiter.getEquation().add(), IOfficeMath.class);
		// Adds a math script.
		mathScript = addMathScript(officeMath, MathScriptType.Subscript);
		// Adds a math text.
		addMathText(document, mathScript.getEquation(), "a");
		addMathText(document, mathScript.getScript(), "n");
		// Adds a math function.
		IOfficeMathFunction mathFunction = (IOfficeMathFunction) ObjectSupport
				.instanceOf(officeMath.getFunctions().add(MathFunctionType.Function), IOfficeMathFunction.class);
		IOfficeMathRunElement mathParaItem = addMathText(document, mathFunction.getFunctionName(), "cos");
		mathParaItem.getMathFormat().setStyle(MathStyleType.Regular);
		// Adds a math fraction.
		IOfficeMathFraction mathFraction = (IOfficeMathFraction) ObjectSupport.instanceOf(
				mathFunction.getEquation().getFunctions().add(MathFunctionType.Fraction), IOfficeMathFraction.class);
		// Unicode value of PI.
		String pi = "\uD835\uDF0B";
		String text = StringSupport.concat(StringSupport.concat("n", pi), "x");
		// Adds a math text.
		addMathText(document, mathFraction.getNumerator(), text);
		addMathText(document, mathFraction.getDenominator(), "L");
		addMathText(document, officeMath, "+");
		// Adds a math script.
		mathScript = addMathScript(officeMath, MathScriptType.Subscript);
		// Adds a math text.
		addMathText(document, mathScript.getEquation(), "b");
		addMathText(document, mathScript.getScript(), "n");
		// Adds a function.
		mathFunction = (IOfficeMathFunction) ObjectSupport
				.instanceOf(officeMath.getFunctions().add(MathFunctionType.Function), IOfficeMathFunction.class);
		mathParaItem = addMathText(document, mathFunction.getFunctionName(), "sin");
		mathParaItem.getMathFormat().setStyle(MathStyleType.Regular);
		// Adds a math fraction element.
		mathFraction = (IOfficeMathFraction) ObjectSupport.instanceOf(
				mathFunction.getEquation().getFunctions().add(MathFunctionType.Fraction), IOfficeMathFraction.class);
		// Adds a math text for numerator.
		addMathText(document, mathFraction.getNumerator(), text);
		// Adds a math text for denominator.
		addMathText(document, mathFraction.getDenominator(), "L");
	}

	/**
	 * Adds a math text
	 * 
	 * @param document   Represents a Word document to add math text
	 * @param officeMath Represents an office math to add math text
	 * @param text       Represents the text to set for math item
	 */
	private static IOfficeMathRunElement addMathText(WordDocument document, IOfficeMath officeMath, String text)
			throws Exception {
		// Adds math text.
		IOfficeMathRunElement officeMathParaItem = (IOfficeMathRunElement) ObjectSupport
				.instanceOf(officeMath.getFunctions().add(MathFunctionType.RunElement), IOfficeMathRunElement.class);
		// Set math text value.
		officeMathParaItem.setItem(new WTextRange(document));
		((WTextRange) ObjectSupport.instanceOf(officeMathParaItem.getItem(), WTextRange.class)).setText(text);
		return officeMathParaItem;
	}

	/**
	 * Adds a math Subscript or Superscript equation
	 * 
	 * @param officeMath     Represents an office math to add math text
	 * @param mathScriptType Represents type of script to be added
	 */
	private static IOfficeMathScript addMathScript(IOfficeMath officeMath, MathScriptType mathScriptType)
			throws Exception {
		IOfficeMathScript mathScript = (IOfficeMathScript) ObjectSupport
				.instanceOf(officeMath.getFunctions().add(MathFunctionType.SubSuperscript), IOfficeMathScript.class);
		// Sets the script type as Subscript or Superscript.
		mathScript.setScriptType(mathScriptType);
		return mathScript;
	}

	/**
	 * Adds a accent equation
	 * 
	 * @param document        Represents a Word document to add math text
	 * @param officeMath      Represents a office math to add accent equation
	 * @param accentCharValue Represents a accent character
	 * @param text            Represents a text for accent equation
	 */
	private static void addMathAccent(WordDocument document, IOfficeMath officeMath, short accentCharValue, String text)
			throws Exception {
		IOfficeMathAccent mathAccent = (IOfficeMathAccent) ObjectSupport
				.instanceOf(officeMath.getFunctions().add(MathFunctionType.Accent), IOfficeMathAccent.class);
		// Sets the accent character from short value.
		mathAccent.setAccentCharacter(convertShortToString(accentCharValue));
		// Adds a math text.
		IOfficeMathRunElement officeMathParaItem = addMathText(document, mathAccent.getEquation(), text);
	}

	/**
	 * Add values in matrix row
	 * 
	 * @param document      Represents a Word document to add matrix
	 * @param mathMatrixRow Represents a matrix row to add values
	 * @param text          Represents a base text value for Subscript and
	 *                      Superscript equation
	 */
	private static void addMatrixRowValues(WordDocument document, IOfficeMathMatrixRow mathMatrixRow, String text)
			throws Exception {
		// Adds arguments for matrix row.
		IOfficeMath officeMath = mathMatrixRow.getArguments().get(0);
		// Adds a Subscript.
		IOfficeMathScript mathScript = addMathScript(officeMath, MathScriptType.Subscript);
		// Adds a math text.
		IOfficeMathRunElement officeMathParaItem = addMathText(document, mathScript.getEquation(), text);
		officeMathParaItem = addMathText(document, mathScript.getScript(), "x");
		// Adds arguments for matrix row.
		officeMath = mathMatrixRow.getArguments().get(1);
		mathScript = addMathScript(officeMath, MathScriptType.Subscript);
		// Adds a math text.
		officeMathParaItem = addMathText(document, mathScript.getEquation(), text);
		officeMathParaItem = addMathText(document, mathScript.getScript(), "y");
		// Adds arguments for matrix row.
		officeMath = mathMatrixRow.getArguments().get(2);
		mathScript = addMathScript(officeMath, MathScriptType.Subscript);
		// Adds a math text.
		officeMathParaItem = addMathText(document, mathScript.getEquation(), text);
		officeMathParaItem = addMathText(document, mathScript.getScript(), "z");
	}
}
