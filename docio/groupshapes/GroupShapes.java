package groupshapes;

import com.syncfusion.docio.*;
import com.syncfusion.javahelper.system.ObjectSupport;
import com.syncfusion.javahelper.system.drawing.*;

public class GroupShapes {
	public static void main(String[] args) throws Exception {
		// Creates a new Word document
		WordDocument document = new WordDocument();
		// Adds new section to the document
		IWSection section = document.addSection();
		// Sets page setup options to section
		WPageSetup pageSetup = section.getPageSetup();
		pageSetup.setOrientation(PageOrientation.Landscape);
		pageSetup.getMargins().setAll(72);
		pageSetup.setPageSize(new SizeFSupport((float) 792f, (float) 612f));
		// Adds new paragraph to the section
		WParagraph paragraph = (WParagraph) section.addParagraph();
		// Creates new group shape
		GroupShape groupShape = new GroupShape(document);
		// Adds group shape to the paragraph
		paragraph.getChildEntities().add(groupShape);
		// Create a RoundedRectangle shape with "Management" text
		createChildShape(AutoShapeType.RoundedRectangle, 324f, (float) 107.7f, (float) 144f, (float) 45f, (float) 0,
				false, false, ColorSupport.fromArgb(50, 48, 142), "Management", groupShape, document);
		// Create a BentUpArrow shape to connect with "Development" shape
		createChildShape(AutoShapeType.BentUpArrow, 177.75f, (float) 176.25f, (float) 210f, (float) 50f, (float) 180,
				false, false, ColorSupport.getWhite(), null, groupShape, document);
		// Create a BentUpArrow shape to connect with "Sales" shape
		createChildShape(AutoShapeType.BentUpArrow, 403.5f, (float) 175.5f, (float) 210f, (float) 50f, (float) 180,
				true, false, ColorSupport.getWhite(), null, groupShape, document);
		// Create a DownArrow shape to connect with "Production" shape
		createChildShape(AutoShapeType.DownArrow, 381f, (float) 153f, (float) 29.25f, (float) 72.5f, (float) 0, false,
				false, ColorSupport.getWhite(), null, groupShape, document);
		// Create a RoundedRectangle shape with "Development" text
		createChildShape(AutoShapeType.RoundedRectangle, 135f, (float) 226.45f, (float) 110f, (float) 40f, (float) 0,
				false, false, ColorSupport.fromArgb(104, 57, 157), "Development", groupShape, document);
		// Create a RoundedRectangle shape with "Production" text
		createChildShape(AutoShapeType.RoundedRectangle, 341f, (float) 226.5f, (float) 110f, (float) 40f, (float) 0,
				false, false, ColorSupport.fromArgb(149, 50, 118), "Production", groupShape, document);
		// Create a RoundedRectangle shape with "Sales" text
		createChildShape(AutoShapeType.RoundedRectangle, 546.75f, (float) 226.5f, (float) 110f, (float) 40f, (float) 0,
				false, false, ColorSupport.fromArgb(179, 63, 62), "Sales", groupShape, document);
		// Create a DownArrow shape to connect with "Software" and "Hardware" shape
		createChildShape(AutoShapeType.DownArrow, 177f, (float) 265.5f, (float) 25.5f, (float) 20.25f, (float) 0, false,
				false, ColorSupport.getWhite(), null, groupShape, document);
		// Create a DownArrow shape to connect with "Series" and "Parts" shape
		createChildShape(AutoShapeType.DownArrow, 383.25f, (float) 265.5f, (float) 25.5f, (float) 20.25f, (float) 0,
				false, false, ColorSupport.getWhite(), null, groupShape, document);
		// Create a DownArrow shape to connect with "North" and "South" shape
		createChildShape(AutoShapeType.DownArrow, 588.75f, (float) 266.25f, (float) 25.5f, (float) 20.25f, (float) 0,
				false, false, ColorSupport.getWhite(), null, groupShape, document);
		// Create a BentUpArrow shape to connect with "Software" shape
		createChildShape(AutoShapeType.BentUpArrow, 129.5f, (float) 286.5f, (float) 60f, (float) 33f, (float) 180,
				false, false, ColorSupport.getWhite(), null, groupShape, document);
		// Create a BentUpArrow shape to connect with "Hardware" shape
		createChildShape(AutoShapeType.BentUpArrow, 190.5f, (float) 286.5f, (float) 60f, (float) 33f, (float) 180, true,
				false, ColorSupport.getWhite(), null, groupShape, document);
		// Create a BentUpArrow shape to connect with "Series" shape
		createChildShape(AutoShapeType.BentUpArrow, 336f, (float) 287.25f, (float) 60f, (float) 33f, (float) 180, false,
				false, ColorSupport.getWhite(), null, groupShape, document);
		// Create a BentUpArrow shape to connect with "Parts" shape
		createChildShape(AutoShapeType.BentUpArrow, 397f, (float) 287.25f, (float) 60f, (float) 33f, (float) 180, true,
				false, ColorSupport.getWhite(), null, groupShape, document);
		// Create a BentUpArrow shape to connect with "North" shape
		createChildShape(AutoShapeType.BentUpArrow, 541.5f, (float) 288f, (float) 60f, (float) 33f, (float) 180, false,
				false, ColorSupport.getWhite(), null, groupShape, document);
		// Create a BentUpArrow shape to connect with "South" shape
		createChildShape(AutoShapeType.BentUpArrow, 602.5f, (float) 288f, (float) 60f, (float) 33f, (float) 180, true,
				false, ColorSupport.getWhite(), null, groupShape, document);
		// Create a RoundedRectangle shape with "Software" text
		createChildShape(AutoShapeType.RoundedRectangle, 93f, (float) 320.25f, (float) 90f, (float) 40f, (float) 0,
				false, false, ColorSupport.fromArgb(23, 187, 189), "Software", groupShape, document);
		// Create a RoundedRectangle shape with "Hardware" text
		createChildShape(AutoShapeType.RoundedRectangle, 197.2f, (float) 320.25f, (float) 90f, (float) 40f, (float) 0,
				false, false, ColorSupport.fromArgb(24, 159, 106), "Hardware", groupShape, document);
		// Create a RoundedRectangle shape with "Series" text
		createChildShape(AutoShapeType.RoundedRectangle, 299.25f, (float) 320.25f, (float) 90f, (float) 40f, (float) 0,
				false, false, ColorSupport.fromArgb(23, 187, 189), "Series", groupShape, document);
		// Create a RoundedRectangle shape with "Parts" text
		createChildShape(AutoShapeType.RoundedRectangle, 404.2f, (float) 320.25f, (float) 90f, (float) 40f, (float) 0,
				false, false, ColorSupport.fromArgb(24, 159, 106), "Parts", groupShape, document);
		// Create a RoundedRectangle shape with "North" text
		createChildShape(AutoShapeType.RoundedRectangle, 505.5f, (float) 321.75f, (float) 90f, (float) 40f, (float) 0,
				false, false, ColorSupport.fromArgb(23, 187, 189), "North", groupShape, document);
		// Create a RoundedRectangle shape with "South" text
		createChildShape(AutoShapeType.RoundedRectangle, 609.7f, (float) 321.75f, (float) 90f, (float) 40f, (float) 0,
				false, false, ColorSupport.fromArgb(24, 159, 106), "South", groupShape, document);
		// Save and close the document.
		document.save("Group Shapes.docx");
		document.close();
		System.out.println("Word document generated successfully.");
	}

	/**
	 * Create a child shape with its specified properties and add into specified
	 * group shape
	 * 
	 * @param autoShapeType Represent the AutoShapeType of child shape
	 * @param X             Represent horizontal position of shape
	 * @param Y             Represent vertical position of the shape
	 * @param width         Represent height of the shape
	 * @param height        Represent height of the shape
	 * @param rotation      Represent the rotation of child shape
	 * @param flipH         Represent the horizontal flip of child shape
	 * @param flipV         Represent the vertical flip of child shape
	 * @param fillColor     Represent the fill color of child shape
	 * @param text          Represent the text that to be append in child shape
	 * @param groupShape    Represent the group shape to add a child shape
	 * @param wordDocument  Represent the Word document instance
	 * 
	 */

	private static void createChildShape(AutoShapeType autoShapeType, float X, float Y, float width, float height,
			float rotation, boolean flipH, boolean flipV, ColorSupport fillColor, String text, GroupShape groupShape,
			WordDocument wordDocument) throws Exception {
		Shape shape = new Shape(wordDocument, autoShapeType);
		shape.setHeight(height);
		shape.setWidth(width);
		shape.setHorizontalPosition(X);
		shape.setVerticalPosition(Y);
		if (rotation != 0)
			shape.setRotation(rotation);
		if (flipH)
			shape.setFlipHorizontal(true);
		if (flipV)
			shape.setFlipVertical(true);
		if (fillColor != ColorSupport.getWhite()) {
			shape.getFillFormat().setFill(true);
			shape.getFillFormat().setColor(fillColor);
		}
		shape.getWrapFormat().setTextWrappingStyle(TextWrappingStyle.InFrontOfText);
		shape.setHorizontalOrigin(HorizontalOrigin.Page);
		shape.setVerticalOrigin(VerticalOrigin.Page);
		if (autoShapeType == AutoShapeType.RoundedRectangle)
			shape.getLineFormat().setLine(false);
		if (text != null) {
			IWParagraph paragraph = shape.getTextBody().addParagraph();
			shape.getTextFrame().setTextVerticalAlignment(VerticalAlignment.Middle);
			paragraph.getParagraphFormat().setHorizontalAlignment(HorizontalAlignment.Center);
			IWTextRange textRange = paragraph.appendText(text);
			// Applying character formats to the text
			WCharacterFormat characterFormat = textRange.getCharacterFormat();
			characterFormat.setFontName("Calibri");
			characterFormat.setFontSize(15);
			characterFormat.setTextColor(ColorSupport.getWhite());
			characterFormat.setBold(true);
			characterFormat.setItalic(true);
		}
		groupShape.add(shape);
	}
}
