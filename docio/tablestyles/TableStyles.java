package tablestyles;

import java.io.*;
import com.syncfusion.docio.*;
import com.syncfusion.javahelper.system.*;
import com.syncfusion.javahelper.system.collections.generic.ListSupport;
import com.syncfusion.javahelper.system.drawing.*;
import com.syncfusion.javahelper.system.io.*;
import com.syncfusion.javahelper.system.xml.*;;

public class TableStyles {
	public static void main(String[] args) throws Exception {
		// Creates a new document.
		WordDocument document = new WordDocument();
		// Loads the template document.
		document.open(getDataDir("TemplateTableStyle.docx"), FormatType.Docx);
		// Creates MailMergeDataTable.
		MailMergeDataTable mailMergeDataTable = getMailMergeDataTable();
		// Executes Mail Merge with groups.
		document.getMailMerge().executeGroup(mailMergeDataTable);
		// Creates table and apply table styles.
		WTable table = createAndApplyTableStyles(document);
		// Saves and closes the document.
		document.save("Table Styles.docx", FormatType.Docx);
		System.out.println("Word document generated successfully.");
	}

	/**
	 * Creates table and apply table styles
	 * @param document Represents a Word document instance
	 * 
	 */
	static WTable createAndApplyTableStyles(WordDocument document) throws Exception {
		WTable table = (WTable) document.getLastSection().getTables().get(0);
		WTableStyle tableStyle = (WTableStyle)document.addTableStyle("Tablestyle");				
		// Applies Table style.
		ColorSupport borderColor = (ColorSupport.getWhiteSmoke()).clone();
		ColorSupport firstRowBackColor = (ColorSupport.getBlue()).clone();
		ColorSupport backColor = (ColorSupport.getWhiteSmoke()).clone();
		ConditionalFormattingStyle firstRowStyle, lastRowStyle, firstColumnStyle, lastColumnStyle,
				oddColumnBandingStyle, oddRowBandingStyle, evenRowBandingStyle;
		// Applies Table Properties.
		TableStyleTableProperties tableProperties = tableStyle.getTableProperties();
		tableProperties.setRowStripe(1);
		tableProperties.setColumnStripe(1);
		tableProperties.setLeftIndent(0);
		// Applies padding properties to table.
		Paddings padding = tableStyle.getTableProperties().getPaddings();
		padding.setTop(0);
		padding.setBottom(0);
		padding.setLeft(5.4f);
		padding.setRight(5.4f);
		
		// Applies border properties to table.
		Border border = tableStyle.getTableProperties().getBorders().getTop();
		applyBorderPropertiesInTable(border,BorderStyle.Single,1f,ColorSupport.getAliceBlue().clone(),0);		
		border = tableStyle.getTableProperties().getBorders().getBottom();
		applyBorderPropertiesInTable(border,BorderStyle.Single,1f,(borderColor).clone(),0);
		border = tableStyle.getTableProperties().getBorders().getLeft();
		applyBorderPropertiesInTable(border,BorderStyle.Single,1f,(borderColor).clone(),0);
		border = tableStyle.getTableProperties().getBorders().getRight();
		applyBorderPropertiesInTable(border,BorderStyle.Single,1f,(borderColor).clone(),0);
		border = tableStyle.getTableProperties().getBorders().getHorizontal();
		border.setBorderType(BorderStyle.Single);
		border.setLineWidth(1f);
		border.setColor((borderColor).clone());
		border = tableStyle.getTableProperties().getBorders().getVertical();
		border.setSpace(0);
		// Adds formatting styles for first row.
		firstRowStyle = tableStyle.getConditionalFormattingStyles().add(ConditionalFormattingType.FirstRow);
		// Applies character format.
		WCharacterFormat characterFormat = firstRowStyle.getCharacterFormat();
		characterFormat.setBold(true);
		characterFormat.setBoldBidi(true);
		characterFormat.setTextColor((ColorSupport.fromArgb(255, 255, 255, 255)).clone());
		
		// Applies border properties to firstRow.
		border = firstRowStyle.getCellProperties().getBorders().getTop();
		applyBorderPropertiesInTable(border,BorderStyle.Single,1f,(borderColor).clone(),0);
		border = firstRowStyle.getCellProperties().getBorders().getBottom();
		applyBorderPropertiesInTable(border,BorderStyle.Single,1f,(borderColor).clone(),0);
		border = firstRowStyle.getCellProperties().getBorders().getLeft();
		applyBorderPropertiesInTable(border,BorderStyle.Single,1f,(borderColor).clone(),0);
		border = firstRowStyle.getCellProperties().getBorders().getRight();
		applyBorderPropertiesInTable(border,BorderStyle.Single,1f,(borderColor).clone(),0);
		border = firstRowStyle.getCellProperties().getBorders().getHorizontal();
		border.setBorderType(BorderStyle.Cleared);
		border.setBorderType(BorderStyle.Cleared);
		
		// Applies table cell properties.
		TableStyleCellProperties cellProperties = firstRowStyle.getCellProperties();
		cellProperties.setBackColor((firstRowBackColor).clone());
		cellProperties.setForeColor((ColorSupport.fromArgb(0, 255, 255, 255)).clone());
		cellProperties.setTextureStyle(TextureStyle.TextureNone);
		lastRowStyle = tableStyle.getConditionalFormattingStyles().add(ConditionalFormattingType.LastRow);
		characterFormat = lastRowStyle.getCharacterFormat();
		characterFormat.setBold(true);
		characterFormat.setBoldBidi(true);
		
		// Applies border properties to lastRow.
		border = lastRowStyle.getCellProperties().getBorders().getTop();
		applyBorderPropertiesInTable(border,BorderStyle.Double,.75f,(borderColor).clone(),0);
		border = lastRowStyle.getCellProperties().getBorders().getBottom();
		applyBorderPropertiesInTable(border,BorderStyle.Single,1f,(borderColor).clone(),0);
		border = lastRowStyle.getCellProperties().getBorders().getLeft();
		applyBorderPropertiesInTable(border,BorderStyle.Single,1f,(borderColor).clone(),0);
		border = lastRowStyle.getCellProperties().getBorders().getRight();
		applyBorderPropertiesInTable(border,BorderStyle.Single,1f,(borderColor).clone(),0);
		lastRowStyle.getCellProperties().getBorders().getHorizontal().setBorderType(BorderStyle.Cleared);
		lastRowStyle.getCellProperties().getBorders().getVertical().setBorderType(BorderStyle.Cleared);
		firstColumnStyle = tableStyle.getConditionalFormattingStyles().add(ConditionalFormattingType.FirstColumn);
		
		// Applies Character format.
		characterFormat = firstColumnStyle.getCharacterFormat();
		characterFormat.setBold(true);
		characterFormat.setBoldBidi(true);
		lastColumnStyle = tableStyle.getConditionalFormattingStyles().add(ConditionalFormattingType.LastColumn);
		characterFormat = lastColumnStyle.getCharacterFormat();
		characterFormat.setBold(true);
		characterFormat.setBoldBidi(true);
		oddColumnBandingStyle = tableStyle.getConditionalFormattingStyles()
				.add(ConditionalFormattingType.OddColumnBanding);
		
		// Applies Table Cell Properties.
		cellProperties = oddColumnBandingStyle.getCellProperties();
		cellProperties.setBackColor((backColor).clone());
		cellProperties.setForeColor((ColorSupport.fromArgb(0, 255, 255, 255)).clone());
		cellProperties.setTextureStyle(TextureStyle.TextureNone);
		oddRowBandingStyle = tableStyle.getConditionalFormattingStyles().add(ConditionalFormattingType.OddRowBanding);
		
		// Applies Table Cell Properties.
		cellProperties = oddRowBandingStyle.getCellProperties();
		cellProperties.getBorders().getHorizontal().setBorderType(BorderStyle.Cleared);
		cellProperties.getBorders().getVertical().setBorderType(BorderStyle.Cleared);
		cellProperties.setBackColor((backColor).clone());
		cellProperties.setForeColor((ColorSupport.fromArgb(0, 255, 255, 255)).clone());
		cellProperties.setTextureStyle(TextureStyle.TextureNone);
		evenRowBandingStyle = tableStyle.getConditionalFormattingStyles().add(ConditionalFormattingType.EvenRowBanding);
		cellProperties = evenRowBandingStyle.getCellProperties();
		cellProperties.getBorders().getHorizontal().setBorderType(BorderStyle.Cleared);
		cellProperties.getBorders().getVertical().setBorderType(BorderStyle.Cleared);
		table.applyStyle("Tablestyle");
		return table;
	}

	/**
	 * 
	 * Gets the mail merge data table.
	 * 
	 * @return
	 * @exception System.Exception reader
	 * @exception XmlException     Unexpected xml tag + reader.LocalName
	 */
	private static MailMergeDataTable getMailMergeDataTable() throws Exception {
		// Gets list of suppliers details.
		ListSupport<Suppliers> suppliers = new ListSupport<Suppliers>(Suppliers.class);
		FileStreamSupport fs = new FileStreamSupport(getDataDir("Suppliers.xml"), FileMode.Open, FileAccess.Read);
		// Reads the xml document.
		XmlReaderSupport reader = XmlReaderSupport.create(fs);
		if (reader == null)
			throw new Exception("reader");
		while (reader.getNodeType().getEnumValue() != XmlNodeType.Element.getEnumValue())
			reader.read();
		if (!(reader.getLocalName() == "SuppliersList"))
			throw new Exception("Unexpected xml tag " +  reader.getLocalName());
		reader.read();
		while (reader.getNodeType().getEnumValue() == XmlNodeType.Whitespace.getEnumValue())
			reader.read();
		// Iterates to add the suppliers details in list.
		while (!(reader.getLocalName() == "SuppliersList")) {
			if (reader.getNodeType().getEnumValue() == XmlNodeType.Element.getEnumValue()) {
				switch ((reader.getLocalName()) == null ? "string_null_value" : (reader.getLocalName())) {
				case "Suppliers":

					suppliers.add(getSupplier(reader));
					break;
				}
			} else

			{
				reader.read();
				if ((reader.getLocalName() == "SuppliersList")
						&& reader.getNodeType().getEnumValue() == XmlNodeType.EndElement.getEnumValue())
					break;
			}
		}
		//Creates an instance of “MailMergeDataTable” by specifying mail merge group name and “IEnumerable” collection.
		MailMergeDataTable dataTable = new MailMergeDataTable("Suppliers", suppliers);
		reader.close();
		fs.close();
		return dataTable;
	}

	/**
	 * 
	 * Gets the suppliers.
	 * 
	 * @param reader The reader.
	 * @return
	 * @exception System.Exception reader
	 * @exception XmlException     Unexpected xml tag + reader.LocalName
	 */
	private static Suppliers getSupplier(XmlReaderSupport reader) throws Exception {
		if (reader == null)
			throw new Exception("reader");
		while (reader.getNodeType().getEnumValue() != XmlNodeType.Element.getEnumValue())
			reader.read();
		if (!(reader.getLocalName() == "Suppliers"))
			throw new Exception("Unexpected xml tag " + reader.getLocalName());
		reader.read();
		while (reader.getNodeType().getEnumValue() == XmlNodeType.Whitespace.getEnumValue())
			reader.read();
		Suppliers supplier = new Suppliers();
		while (!(reader.getLocalName() == "Suppliers")) {
			if (reader.getNodeType().getEnumValue() == XmlNodeType.Element.getEnumValue()) {
				switch ((reader.getLocalName()) == null ? "string_null_value" : (reader.getLocalName())) {
				case "SupplierID":

					supplier.setSupplierID(reader.readContentAsString());
					break;
				case "CompanyName":

					supplier.setCompanyName(reader.readContentAsString());
					break;
				case "ContactName":

					supplier.setContactName(reader.readContentAsString());
					break;
				case "ContactTitle":

					supplier.setContactTitle(reader.readContentAsString());
					break;
				case "Region":

					supplier.setRegion(reader.readContentAsString());
					break;
				case "PostalCode":

					supplier.setPostalCode(reader.readContentAsString());
					break;
				case "Phone":

					supplier.setPhone(reader.readContentAsString());
					break;
				case "Fax":

					supplier.setFax(reader.readContentAsString());
					break;
				case "HomePage":

					supplier.setHomePage(reader.readContentAsString());
					break;
				case "Address":

					supplier.setAddress(reader.readContentAsString());
					break;
				case "City":

					supplier.setCity(reader.readContentAsString());
					break;
				case "Country":

					supplier.setCountry(reader.readContentAsString());
					break;
				default:

					reader.skip();
					break;
				}
			} else

			{
				reader.read();
				if ((reader.getLocalName() == "Suppliers")
						&& reader.getNodeType().getEnumValue() == XmlNodeType.EndElement.getEnumValue())
					break;
			}
		}
		return supplier;
	}
	/**
	 * Apply border properties to table.
	 * @param borderStyles
	 * @param styles
	 * @param lineWidth
	 * @param color
	 * @param space
	 * @throws Exception
	 */
	static Border applyBorderPropertiesInTable(Border borderStyles, BorderStyle styles, float lineWidth, ColorSupport color, float space ) throws Exception {
		borderStyles.setBorderType(styles);
		borderStyles.setLineWidth(lineWidth);
		borderStyles.setColor(color);
		borderStyles.setSpace(space);
		return borderStyles;
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
