package autoshapes;

import java.io.*;
import com.syncfusion.docio.*;
import com.syncfusion.javahelper.system.drawing.*;

public class AutoShapes {
	public static void main(String[] args) throws Exception {
		// Initialize Word document.
		WordDocument doc = new WordDocument();
		// Ensure Minimum.
		doc.ensureMinimal();
		// Create AutoShape.
		Shape shape = createShape(doc, AutoShapeType.RoundedRectangle, (float) 130, (float) 45, (float) 50, true,
				ColorSupport.getBlue(), VerticalAlignment.Middle);
		// Add Texbody contents to Shape.
		IWParagraph para = addParagraph(shape);
		IWTextRange textRange = appendText(para, "Requirement");
		shape = createDownArrow(doc, AutoShapeType.DownArrow, (float) 45, (float) 45, (float) 95, true);
		shape = createShape(doc, AutoShapeType.RoundedRectangle, (float) 130, (float) 45, (float) 140, true,
				ColorSupport.getOrange(), VerticalAlignment.Middle);
		// Add Texbody contents to Shape.
		para = addParagraph(shape);
		textRange = appendText(para, "Design");
		shape = createDownArrow(doc, AutoShapeType.DownArrow, (float) 45, (float) 45, (float) 185, true);
		shape = createShape(doc, AutoShapeType.RoundedRectangle, (float) 130, (float) 45, (float) 230, true,
				ColorSupport.getBlue(), VerticalAlignment.Middle);
		// Add Texbody contents to Shape.
		para = addParagraph(shape);
		textRange = appendText(para, "Execution");
		shape = createDownArrow(doc, AutoShapeType.DownArrow, (float) 45, (float) 45, (float) 275, true);
		shape = createShape(doc, AutoShapeType.RoundedRectangle, (float) 130, (float) 45, (float) 320, true,
				ColorSupport.getViolet(), VerticalAlignment.Middle);
		// Add Texbody contents to Shape.
		para = addParagraph(shape);
		textRange = appendText(para, "Testing");
		shape = createDownArrow(doc, AutoShapeType.DownArrow, (float) 45, (float) 45, (float) 365, true);
		shape = createShape(doc, AutoShapeType.RoundedRectangle, (float) 130, (float) 45, (float) 410, true,
				ColorSupport.getPaleVioletRed(), VerticalAlignment.Middle);
		// Add Texbody contents to Shape.
		para = addParagraph(shape);
		textRange = appendText(para, "Release");
		// Save and close the document.
		doc.save("Auto Shapes.docx");
		doc.close();
		System.out.println("Word document generated successfully.");
	}

	/**
	 * Create AutoShape.
	 * 
	 * @param autoShapeType     Represents the shape type.
	 * @param width             Represents the width of the shape.
	 * @param height            Represents the height of the shape.
	 * @param verticalPosition  Represents vertical position of the shape.
	 * @param isAllowOverlap    Represents given shape can overlap other shapes.
	 * @param color             Represents color to fill.
	 * @param verticalAlignment Represents color to fill.
	 */

	private static Shape createShape(WordDocument doc, AutoShapeType autoShapeType, float width, float height,
			float verticalPosition, boolean isAllowOverlap, ColorSupport color, VerticalAlignment verticalAlignment)
			throws Exception {
		// Append AutoShape.
		Shape shape = doc.getLastParagraph().appendShape(autoShapeType, width, height);
		// Set horizontal alignment.
		shape.setHorizontalAlignment(ShapeHorizontalAlignment.Center);
		// Set horizontal origin.
		shape.setHorizontalOrigin(HorizontalOrigin.Page);
		// Set vertical origin.
		shape.setVerticalOrigin(VerticalOrigin.Page);
		// Set vertical position.
		shape.setVerticalPosition(verticalPosition);
		// Set AllowOverlap to true for overlapping shapes.
		shape.getWrapFormat().setAllowOverlap(isAllowOverlap);
		// Set Fill Color.
		shape.getFillFormat().setColor(color);
		// Set Content vertical alignment.
		shape.getTextFrame().setTextVerticalAlignment(verticalAlignment);
		return shape;
	}

	/**
	 * Create AutoShape.
	 * 
	 * @param autoShapeType    Represents the shape type.
	 * @param width            Represents the width of the shape.
	 * @param height           Represents the height of the shape.
	 * @param verticalPosition Represents vertical position of the shape.
	 * @param isAllowOverlap   Represents given shape can overlap other shapes.
	 * 
	 */
	private static Shape createDownArrow(WordDocument doc, AutoShapeType autoShapeType, float width, float height,
			float verticalPosition, boolean isAllowOverlap) throws Exception {
		// Append AutoShape.
		Shape shape = doc.getLastParagraph().appendShape(autoShapeType, width, height);
		// Set horizontal alignment.
		shape.setHorizontalAlignment(ShapeHorizontalAlignment.Center);
		// Set horizontal origin.
		shape.setHorizontalOrigin(HorizontalOrigin.Page);
		// Set vertical origin.
		shape.setVerticalOrigin(VerticalOrigin.Page);
		// Set vertical position.
		shape.setVerticalPosition(verticalPosition);
		// Set AllowOverlap to true for overlapping shapes.
		shape.getWrapFormat().setAllowOverlap(isAllowOverlap);
		return shape;
	}

	/**
	 * 
	 * Add the paragraph
	 * 
	 * @param shape Represents the Shape
	 * 
	 */

	private static IWParagraph addParagraph(Shape shape) throws Exception {
		IWParagraph para = shape.getTextBody().addParagraph();
		para.getParagraphFormat().setHorizontalAlignment(HorizontalAlignment.Center);
		return para;
	}

	/**
	 * 
	 * Add texbody contents to Shape.
	 * 
	 * @param para specifies the paragraph in which the text is appended.
	 * @param text specifies the textrange to be added
	 * 
	 */
	private static IWTextRange appendText(IWParagraph para, String text) throws Exception {
		// Add text contents to paragraph.
		IWTextRange textRange = para.appendText(text);
		// Apply the character formats.
		WCharacterFormat characterFormat = textRange.getCharacterFormat();
		characterFormat.setBold(true);
		characterFormat.setTextColor(ColorSupport.getWhite());
		characterFormat.setFontSize(12f);
		characterFormat.setFontName("Verdana");
		return textRange;
	}
}
