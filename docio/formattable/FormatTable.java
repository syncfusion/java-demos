package formattable;

import java.io.*;
import com.syncfusion.docio.*;
import com.syncfusion.javahelper.system.*;
import com.syncfusion.javahelper.system.drawing.*;

public class FormatTable {

	public static void main(String[] args) throws Exception {
		// Creates a new document.
		WordDocument document = new WordDocument();
		// Adds a new section to the document.
		IWSection section = document.addSection();
		// Sets page setup to the section.
		WPageSetup pageSetup = section.getPageSetup();
		pageSetup.getMargins().setAll((float) 72);
		pageSetup.setDifferentFirstPage(true);
		// Creates and apply formatting for table
		applyTableFormatting(document, section);
		// Saves and closes the document.
		document.save("Format Table.docx");
		document.close();
		System.out.println("Word document generated successfully.");
	}

	/**
	 * Apply formatting for table
	 * 
	 * @param document Represents the word document instance.
	 * @param section  Represents a section.
	 * @return
	 * @throws Exception
	 */
	static WordDocument applyTableFormatting(WordDocument document, IWSection section) throws Exception {
		IWTextRange textRange;
		// Adds paragraph to the section.
		IWParagraph paragraph = section.addParagraph();
		IWParagraph hParagraph = new WParagraph(document);
		// Appends text to the paragraph.
		hParagraph.appendText("Header text\r\n").getCharacterFormat().setFontSize((float) 14);
		section.getHeadersFooters().getFirstPageHeader().getParagraphs().add(hParagraph);

		// Inserts table into the document.
		IWTable hTable = document.getLastSection().getHeadersFooters().getFirstPageHeader().addTable();
		hTable.resetCells(2, 2);
		appendTextToTableCell(hTable, 0, 0, "1");
		appendTextToTableCell(hTable, 0, 1, "2");
		appendTextToTableCell(hTable, 1, 0, "3");
		appendTextToTableCell(hTable, 1, 1, "4");

		// Adds paragraph to the section.
		paragraph = section.addParagraph();
		paragraph.appendText("Tiny table\r\n").getCharacterFormat().setFontSize((float) 14);

		// Adds paragraph to the section.
		paragraph = section.addParagraph();
		WTextBody textBody = section.getBody();

		// Inserts table.
		IWTable table = textBody.addTable();
		table.resetCells(2, 2);
		appendTextToTableCell(table, 0, 0, "A");
		appendTextToTableCell(table, 0, 0, "AA");
		appendTextToTableCell(table, 0, 0, "AAA");
		appendTextToTableCell(table, 1, 1, "B");
		appendTextToTableCell(table, 1, 1, "BB\\r\\nBBB");
		appendTextToTableCell(table, 1, 1, "BBB");
		textBody.addParagraph().appendText("Text after table...").getCharacterFormat().setFontSize((float) 14);

		// Adds paragraph to the section.
		section.addParagraph();
		paragraph = section.addParagraph();
		// Table with different formatting.
		paragraph.appendText("Table with different formatting\r\n").getCharacterFormat().setFontSize((float) 14);
		// Adds paragraph to the section.
		paragraph = section.addParagraph();
		textBody = section.getBody();

		// Inserts table.
		table = textBody.addTable();
		table.resetCells(3, 3);
		// Gets first row from table.
		WTableRow row0 = table.getRows().get(0);
		paragraph = (IWParagraph) row0.getCells().get(0).addParagraph();
		// Add text to the paragraph.
		textRange = paragraph.appendText("1");
		paragraph.getParagraphFormat().setHorizontalAlignment(HorizontalAlignment.Left);
		// Applies formats to text.
		WCharacterFormat characterFormat = textRange.getCharacterFormat();
		characterFormat.setFontName("Arial");
		characterFormat.setBold(true);
		characterFormat.setFontSize(14f);
		// Applies border properties to cell.
		Borders borders = row0.getCells().get(0).getCellFormat().getBorders();
		borders.setLineWidth(2f);
		borders.setColor((ColorSupport.getMagenta()).clone());
		// Adds paragraph in firstrow of secondcell.
		paragraph = (IWParagraph) row0.getCells().get(1).addParagraph();
		textRange = paragraph.appendText("2");
		paragraph.getParagraphFormat().setHorizontalAlignment(HorizontalAlignment.Right);
		// Applies character formats to text.
		characterFormat = textRange.getCharacterFormat();
		characterFormat.setEmboss(true);
		characterFormat.setFontSize(15f);
		// Applies border properties to first row of second cell.
		borders = row0.getCells().get(1).getCellFormat().getBorders();
		borders.setLineWidth(1.3f);
		borders.setBorderType(BorderStyle.DoubleWave);
		// Adds paragraph to firstrow of thirdcell.
		paragraph = (IWParagraph) row0.getCells().get(2).addParagraph();
		textRange = paragraph.appendText("3");
		paragraph.getParagraphFormat().setHorizontalAlignment(HorizontalAlignment.Left);
		// Applies character formats to text.
		characterFormat = textRange.getCharacterFormat();
		characterFormat.setEngrave(true);
		characterFormat.setFontSize(15f);
		row0.getCells().get(2).getCellFormat().getBorders().setBorderType(BorderStyle.Emboss3D);
		// Gets second row from table.
		WTableRow row1 = table.getRows().get(1);
		paragraph = (IWParagraph) row1.getCells().get(0).addParagraph();
		textRange = paragraph.appendText("4");
		paragraph.getParagraphFormat().setHorizontalAlignment(HorizontalAlignment.Center);
		characterFormat = textRange.getCharacterFormat();
		characterFormat.setSmallCaps(true);
		characterFormat.setFontName("Comic Sans MS");
		characterFormat.setFontSize((float) 16);
		// Applies border properties to second row of first cell.
		borders = row1.getCells().get(0).getCellFormat().getBorders();
		borders.setLineWidth(2f);
		borders.setBorderType(BorderStyle.DashDotStroker);
		paragraph = (IWParagraph) row1.getCells().get(1).addParagraph();
		textRange = paragraph.appendText("5");
		paragraph.getParagraphFormat().setHorizontalAlignment(HorizontalAlignment.Center);
		// Applies character formats to text.
		characterFormat = textRange.getCharacterFormat();
		characterFormat.setFontName("Times New Roman");
		characterFormat.setShadow(true);
		characterFormat.setTextBackgroundColor((ColorSupport.getOrange()).clone());
		characterFormat.setFontSize(15f);
		// Applies border properties to second row of second cell.
		borders = row1.getCells().get(1).getCellFormat().getBorders();
		borders.setLineWidth(2f);
		borders.setColor((ColorSupport.getBrown()).clone());
		paragraph = (IWParagraph) row1.getCells().get(2).addParagraph();
		textRange = paragraph.appendText("6");
		paragraph.getParagraphFormat().setHorizontalAlignment(HorizontalAlignment.Center);
		// Applies character formats to text.
		characterFormat = textRange.getCharacterFormat();
		characterFormat.setBold(true);
		characterFormat.setFontSize(14f);
		// Applies cell format to second row of third cell.
		CellFormat cellFormat = row1.getCells().get(2).getCellFormat();
		cellFormat.setBackColor((ColorSupport.fromArgb(51, 51, 101)).clone());
		cellFormat.setVerticalAlignment(VerticalAlignment.Middle);
		// Gets third row from table.
		WTableRow row2 = table.getRows().get(2);
		paragraph = (IWParagraph) row2.getCells().get(0).addParagraph();
		textRange = paragraph.appendText("7");
		paragraph.getParagraphFormat().setHorizontalAlignment(HorizontalAlignment.Right);
		textRange.getCharacterFormat().setFontSize(13f);
		// Applies border properties to third row of first cell.
		borders = row2.getCells().get(0).getCellFormat().getBorders();
		borders.setLineWidth(1.5f);
		borders.setBorderType(BorderStyle.DashLargeGap);
		paragraph = (IWParagraph) row2.getCells().get(1).addParagraph();
		textRange = paragraph.appendText("8");
		// Applies character formats to text.
		characterFormat = textRange.getCharacterFormat();
		characterFormat.setTextColor((ColorSupport.getBlue()).clone());
		characterFormat.setFontSize(16f);
		// Applies border properties to third row of second cell.
		borders = row2.getCells().get(1).getCellFormat().getBorders();
		borders.setLineWidth(3f);
		borders.setBorderType(BorderStyle.Wave);
		paragraph = (IWParagraph) row2.getCells().get(2).addParagraph();
		textRange = paragraph.appendText("9");
		paragraph.getParagraphFormat().setHorizontalAlignment(HorizontalAlignment.Right);
		// Applies border properties to third row of third cell.
		borders = row2.getCells().get(2).getCellFormat().getBorders();
		borders.setLineWidth(2f);
		borders.setColor((ColorSupport.getCyan()).clone());
		borders.setShadow(true);
		borders.setSpace((float) 20);
		section.addParagraph();
		// Table Cell Merging.
		paragraph = section.addParagraph();
		paragraph.appendText("Table Cell Merging...").getCharacterFormat().setFontSize((float) 14);
		section.addParagraph();
		paragraph = section.addParagraph();
		textBody = section.getBody();

		// Appends table to text body.
		table = textBody.addTable();
		RowFormat format = new RowFormat();
		format.getPaddings().setAll((float) 5);
		format.getBorders().setBorderType(BorderStyle.Dot);
		format.getBorders().setLineWidth((float) 2);
		table.resetCells(6, 6, format, (float) 80);
		// Applies cell format to first row of first cell.
		cellFormat = table.getRows().get(0).getCells().get(0).getCellFormat();
		cellFormat.setHorizontalMerge(CellMerge.Start);
		table.getRows().get(0).getCells().get(1).getCellFormat().setHorizontalMerge(CellMerge.Continue);
		cellFormat.setVerticalAlignment(VerticalAlignment.Middle);
		cellFormat.setBackColor((ColorSupport.fromArgb(218, 230, 246)).clone());
		IWParagraph par = table.getRows().get(0).getCells().get(0).addParagraph();
		par.getParagraphFormat().setHorizontalAlignment(HorizontalAlignment.Center);
		par.appendText("Horizontal Merge").getCharacterFormat().setBold(true);
		// Applies cell format to second row of fourth cell.
		cellFormat = table.getRows().get(2).getCells().get(3).getCellFormat();
		cellFormat.setVerticalMerge(CellMerge.Start);
		table.getRows().get(3).getCells().get(3).getCellFormat().setVerticalMerge(CellMerge.Continue);
		cellFormat.setVerticalAlignment(VerticalAlignment.Middle);
		par = table.getRows().get(2).getCells().get(3).addParagraph();
		cellFormat.setBackColor((ColorSupport.fromArgb(252, 172, 85)).clone());
		par.getParagraphFormat().setHorizontalAlignment(HorizontalAlignment.Center);
		par.appendText("Vertical Merge").getCharacterFormat().setBold(true);
		// Table Cell spacing.
		section.addParagraph();
		paragraph = section.addParagraph();
		paragraph.appendText("Table Cell spacing...").getCharacterFormat().setFontSize((float) 14);
		section.addParagraph();
		paragraph = section.addParagraph();
		textBody = section.getBody();

		// Appends table to text body.
		table = textBody.addTable();
		format = new RowFormat();
		format.setCellSpacing((float) 5);
		format.getPaddings().setAll((float) 5);
		format.setCellSpacing(2.5f);
		format.getBorders().setBorderType(BorderStyle.DotDash);
		format.setIsBreakAcrossPages(true);
		table.resetCells(25, 5, format, (float) 80);
		IWTextRange text;
		table.getRows().get(0).setIsHeader(true);
		// Adds paragraph to cells.
		for (int i = 0; i < table.getRows().get(0).getCells().getCount(); i++) {
			paragraph = (WParagraph) table.get(0, i).addParagraph();
			paragraph.getParagraphFormat().setHorizontalAlignment(HorizontalAlignment.Center);
			text = paragraph.appendText(StringSupport.format("Header {0}", i + 1));
			characterFormat = text.getCharacterFormat();
			// Applies character format to text.
			characterFormat.setFont(new FontSupport("Bitstream Vera Serif", 10));
			characterFormat.setBold(true);
			characterFormat.setTextColor((ColorSupport.fromArgb(0, 21, 84)).clone());
			table.get(0, i).getCellFormat().setBackColor((ColorSupport.fromArgb(203, 211, 226)).clone());
		}
		for (int i = 1; i < table.getRows().getCount(); i++) {
			for (int j = 0; j < 5; j++) {
				paragraph = (WParagraph) table.get(i, j).addParagraph();
				paragraph.getParagraphFormat().setHorizontalAlignment(HorizontalAlignment.Center);
				text = paragraph.appendText(StringSupport.format("Cell {0} , {1}", i, j + 1));
				// Applies character format to text.
				characterFormat = text.getCharacterFormat();
				characterFormat.setTextColor((ColorSupport.fromArgb(242, 151, 50)).clone());
				characterFormat.setBold(true);
				if (i % 2 != 1)
					table.get(i, j).getCellFormat().setBackColor((ColorSupport.fromArgb(231, 235, 245)).clone());
				else
					table.get(i, j).getCellFormat().setBackColor((ColorSupport.fromArgb(246, 249, 255)).clone());
			}
		}

		// Nested Table.
		paragraph = section.addParagraph();
		paragraph.appendText("Nested Table...").getCharacterFormat().setFontSize((float) 14);
		section.addParagraph();
		paragraph = section.addParagraph();
		textBody = section.getBody();
		table = textBody.addTable();
		format = new RowFormat();
		format.getPaddings().setAll((float) 5);
		format.setCellSpacing(2.5f);
		format.getBorders().setBorderType(BorderStyle.DotDash);
		table.resetCells(5, 3, format, (float) 100);
		table.resetCells(5, 3, format, (float) 100);
		for (int i = 0; i < table.getRows().get(0).getCells().getCount(); i++) {
			paragraph = (WParagraph) table.get(0, i).addParagraph();
			paragraph.getParagraphFormat().setHorizontalAlignment(HorizontalAlignment.Center);
			text = paragraph.appendText(StringSupport.format("Header {0}", i + 1));
			// Applies character format to text.
			characterFormat = text.getCharacterFormat();
			characterFormat.setFont(new FontSupport("Bitstream Vera Serif", 10));
			characterFormat.setBold(true);
			characterFormat.setTextColor((ColorSupport.fromArgb(0, 21, 84)).clone());
			table.get(0, i).getCellFormat().setBackColor((ColorSupport.fromArgb(242, 151, 50)).clone());
		}
		table.get(0, 2).setWidth((float) 200);
		for (int i = 1; i < table.getRows().getCount(); i++) {
			for (int j = 0; j < 3; j++) {
				paragraph = (WParagraph) table.get(i, j).addParagraph();
				paragraph.getParagraphFormat().setHorizontalAlignment(HorizontalAlignment.Center);
				if ((i == 2) && (j == 2)) {
					text = paragraph.appendText("Nested Table");
				} else

				{
					text = paragraph.appendText(StringSupport.format("Cell {0} , {1}", i, j + 1));
				}
				if ((j == 2))
					table.get(i, j).setWidth((float) 200);
				// Applies character format to text.
				characterFormat = text.getCharacterFormat();
				characterFormat.setTextColor((ColorSupport.fromArgb(242, 151, 50)).clone());
				characterFormat.setBold(true);
			}
		}
		IWTable nestTable = table.get(2, 2).addTable();
		format = new RowFormat();
		format.getBorders().setBorderType(BorderStyle.DotDash);
		format.setHorizontalAlignment(RowAlignment.Center);
		nestTable.resetCells(3, 3, format, (float) 50);
		// Adds paragraph to the nested table.
		for (int i = 0; i < nestTable.getRows().getCount(); i++) {
			for (int j = 0; j < 3; j++) {
				paragraph = (WParagraph) nestTable.get(i, j).addParagraph();
				paragraph.getParagraphFormat().setHorizontalAlignment(HorizontalAlignment.Center);
				nestTable.get(i, j).getCellFormat().setBackColor((ColorSupport.fromArgb(231, 235, 245)).clone());
				text = paragraph.appendText(StringSupport.format("Cell {0} , {1}", i, j + 1));
				characterFormat = text.getCharacterFormat();
				characterFormat.setTextColor((ColorSupport.getBlack()).clone());
				characterFormat.setBold(true);
			}
		}

		// Table with Images.
		section = document.addSection();
		paragraph = section.addParagraph();
		textRange = paragraph.appendText("Table with Images");
		characterFormat = textRange.getCharacterFormat();
		characterFormat.setFontSize(13f);
		characterFormat.setTextColor((ColorSupport.getDarkBlue()).clone());
		characterFormat.setBold(true);
		section.addParagraph();
		paragraph = section.addParagraph();
		text = null;
		table = section.getBody().addTable();
		table.resetCells(1, 3);
		WTableRow row = table.getRows().get(0);
		row.setHeight(25f);
		for (int i = 0; i < 3; i++) {
			paragraph = (IWParagraph) row.getCells().get(i).addParagraph();
			paragraph.getParagraphFormat().setHorizontalAlignment(HorizontalAlignment.Center);
			switch (i) {
			case 0:
				text = paragraph.appendText("SNO");
				row.getCells().get(i).setWidth(50f);
				break;
			case 1:
				text = paragraph.appendText("Product Name");
				break;
			case 2:
				text = paragraph.appendText("Product Image");
				row.getCells().get(i).setWidth(200f);
				break;
			}
			characterFormat = text.getCharacterFormat();
			characterFormat.setBold(true);
			characterFormat.setFontName("Cambria");
			characterFormat.setFontSize(11f);
			characterFormat.setTextColor((ColorSupport.getWhite()).clone());
			// Applies cell format.
			cellFormat = row.getCells().get(i).getCellFormat();
			cellFormat.setVerticalAlignment(VerticalAlignment.Middle);
			cellFormat.setBackColor((ColorSupport.fromArgb(157, 161, 190)).clone());
			cellFormat.getBorders().setBorderType(BorderStyle.Single);
		}
		int sno = 1;
		row1 = table.addRow(false);
		paragraph = (IWParagraph) row1.getCells().get(0).addParagraph();
		paragraph.getParagraphFormat().setHorizontalAlignment(HorizontalAlignment.Center);
		text = paragraph.appendText(Int32Support.toString(sno));
		characterFormat = text.getCharacterFormat();
		characterFormat.setBold(true);
		characterFormat.setFontSize(10f);
		cellFormat(row1, 0);
		paragraph = (IWParagraph) row1.getCells().get(1).addParagraph();
		paragraph.getParagraphFormat().setHorizontalAlignment(HorizontalAlignment.Center);
		text = paragraph.appendText("Essential DocIO");
		characterFormat = text.getCharacterFormat();
		characterFormat.setBold(true);
		characterFormat.setFontSize(10f);
		characterFormat.setTextColor((ColorSupport.fromArgb(50, 65, 124)).clone());
		cellFormat(row1, 1);
		paragraph = (IWParagraph) row1.getCells().get(2).addParagraph();
		paragraph.appendPicture(new FileInputStream(getDataDir("DocIO.gif")));
		paragraph.getParagraphFormat().setHorizontalAlignment(HorizontalAlignment.Center);
		cellFormat(row1, 2);
		sno++;
		row1 = table.addRow(false);
		paragraph = (IWParagraph) row1.getCells().get(0).addParagraph();
		paragraph.getParagraphFormat().setHorizontalAlignment(HorizontalAlignment.Center);
		text = paragraph.appendText(Int32Support.toString(sno));
		characterFormat = text.getCharacterFormat();
		characterFormat.setBold(true);
		characterFormat.setFontSize(10f);
		cellFormat(row1, 0);
		paragraph = (IWParagraph) row1.getCells().get(1).addParagraph();
		paragraph.getParagraphFormat().setHorizontalAlignment(HorizontalAlignment.Center);
		text = paragraph.appendText("Essential PDF");
		characterFormat = text.getCharacterFormat();
		characterFormat.setBold(true);
		characterFormat.setFontSize(10f);
		characterFormat.setTextColor((ColorSupport.fromArgb(50, 65, 124)).clone());
		cellFormat(row1, 1);
		paragraph = (IWParagraph) row1.getCells().get(2).addParagraph();
		paragraph.appendPicture(new FileInputStream(getDataDir("PDF.gif")));
		paragraph.getParagraphFormat().setHorizontalAlignment(HorizontalAlignment.Center);
		cellFormat(row1, 2);
		sno++;
		row1 = table.addRow(false);
		paragraph = (IWParagraph) row1.getCells().get(0).addParagraph();
		paragraph.getParagraphFormat().setHorizontalAlignment(HorizontalAlignment.Center);
		text = paragraph.appendText(Int32Support.toString(sno));
		characterFormat = text.getCharacterFormat();
		characterFormat.setBold(true);
		characterFormat.setFontSize(10f);
		cellFormat(row1, 0);
		paragraph = (IWParagraph) row1.getCells().get(1).addParagraph();
		paragraph.getParagraphFormat().setHorizontalAlignment(HorizontalAlignment.Center);
		text = paragraph.appendText("Essential XlsIO");
		characterFormat = text.getCharacterFormat();
		characterFormat.setBold(true);
		characterFormat.setFontSize(10f);
		characterFormat.setTextColor((ColorSupport.fromArgb(50, 65, 124)).clone());
		cellFormat(row1, 1);
		paragraph = (IWParagraph) row1.getCells().get(2).addParagraph();
		paragraph.appendPicture(new FileInputStream(getDataDir("XlsIO.gif")));
		paragraph.getParagraphFormat().setHorizontalAlignment(HorizontalAlignment.Center);
		cellFormat(row1, 2);
		sno++;
		return document;
	}

	/**
	 * 
	 * @param row1  Represents a table row to be formatted
	 * @param index Represents the cell index
	 * @return
	 * @throws Exception
	 */
	static WTableRow cellFormat(WTableRow row1, int index) throws Exception {
		row1.getCells().get(index).getCellFormat().setVerticalAlignment(VerticalAlignment.Middle);
		row1.getCells().get(index).getCellFormat().getBorders().setBorderType(BorderStyle.Single);
		row1.getCells().get(index).getCellFormat().setBackColor((ColorSupport.fromArgb(217, 223, 239)).clone());
		return row1;
	}

	/**
	 * Append text to the table cells.
	 * 
	 * @param hTable
	 * @param rowIndex
	 * @param colIndex
	 * @param text
	 * @return
	 * @throws Exception
	 */
	static IWTable appendTextToTableCell(IWTable hTable, int rowIndex, int colIndex, String text) throws Exception {
		hTable.get(rowIndex, colIndex).addParagraph().appendText(text);
		return hTable;
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
