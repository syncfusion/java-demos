package updatefields;


import java.io.File;
import java.util.List;
import javax.xml.bind.*;
import com.syncfusion.docio.*;
import com.syncfusion.javahelper.system.collections.generic.ListSupport;

public class UpdateFields {
	public static void main(String[] args) throws Exception {
		// Loads the template document.
		WordDocument document = new WordDocument(getDataDir("TemplateUpdateFields.docx"), FormatType.Docx);
		// Creates MailMergeDataTable.
		MailMergeDataTable mailMergeDataTableStock = getMailMergeDataTableStock();
		// Executes Mail Merge with groups.
		document.getMailMerge().executeGroup(mailMergeDataTableStock);
		// Updates fields in the document.
		document.updateDocumentFields();
		// Unlinks the fields in Word document.
		unlinkFieldsInDocument(document);
		// Saves and closes the document.
		document.save("Update Fields.docx", FormatType.Docx);
		document.close();
		System.out.println("Word document generated successfully.");
	}

	/**
	 * 
	 * Iterates to document elements and get fields.
	 * 
	 * @param document  Input Word document.
	 * @param fieldType Type of the field to find in document.
	 */
	private static void unlinkFieldsInDocument(WordDocument document) throws Exception {
		// Iterates the section from document to get fields.
		for (Object section_tempObj : document.getSections()) {
			WSection section = (WSection) section_tempObj;
			removeFieldCodesInTextBody(section.getBody());
		}
	}

	/**
	 * 
	 * Iterates into body items.
	 * 
	 * @param textBody Represent a text body of the Word document.
	 */
	private static void removeFieldCodesInTextBody(WTextBody textBody) throws Exception {
		// Iterates the text body item.
		for (int i = 0; i < textBody.getChildEntities().getCount(); i++) {
			IEntity bodyItemEntity = textBody.getChildEntities().get(i);
			switch (bodyItemEntity.getEntityType().toString()) {
			case "Paragraph":

				WParagraph paragraph = (WParagraph)bodyItemEntity;
				removeFieldCodesInParagraph(paragraph.getItems());
				break;
			case "Table":

				removeFieldCodesInTable((WTable)bodyItemEntity);
				break;
			case "BlockContentControl":

				BlockContentControl blockContentControl = (BlockContentControl)bodyItemEntity;						
				removeFieldCodesInTextBody(blockContentControl.getTextBody());
				break;
			}
		}
	}

	/**
	 * 
	 * Iterates into paragraph items.
	 * 
	 * @param paragraph The paragraph.
	 * @param fieldType Type of field.
	 */
	private static void removeFieldCodesInParagraph(ParagraphItemCollection paraItems) throws Exception {
		// Iterates the paragraph items.
		for (int i = 0; i < paraItems.getCount(); i++) {
			if (paraItems.get(i) instanceof WField) {
				WField field = (WField)paraItems.get(i);
				// Unlink the fields.
				field.unlink();
			} else

			if (paraItems.get(i) instanceof WTextBox) {
				WTextBox textBox = (WTextBox)paraItems.get(i);
				removeFieldCodesInTextBody(textBox.getTextBoxBody());
			} else

			if (paraItems.get(i) instanceof Shape) {
				Shape shape = (Shape)paraItems.get(i);
				removeFieldCodesInTextBody(shape.getTextBody());
			} else

			if (paraItems.get(i) instanceof InlineContentControl) {
				InlineContentControl inlineContentControl = (InlineContentControl)paraItems.get(i);
				removeFieldCodesInParagraph(inlineContentControl.getParagraphItems());
			}
		}
	}

	/**
	 * 
	 * Iterates into table.
	 * 
	 * @param table     The table.
	 * @param fieldType Type of Field.
	 */
	private static void removeFieldCodesInTable(WTable table) throws Exception {
		// Iterates the table rows and cells.
		for (Object row_tempObj : table.getRows()) {
			WTableRow row = (WTableRow) row_tempObj;
			for (Object cell_tempObj : row.getCells()) {
				WTableCell cell = (WTableCell) cell_tempObj;
				removeFieldCodesInTextBody(cell);
			}
		}
	}

	/**
	 * 
	 * Gets the mail merge data table.
	 * 
	 */
	private static MailMergeDataTable getMailMergeDataTableStock() throws Exception {
		// Loads the XML file.
		File file = new File(getDataDir("StockDetails.xml"));
		// Create a new instance for the JAXBContext.
		JAXBContext jaxbContext = JAXBContext.newInstance(StockMarket.class);
		// Reads the XML file.
		Unmarshaller jaxbUnmarshaller = jaxbContext.createUnmarshaller();
		StockMarket stockMarket = (StockMarket) jaxbUnmarshaller.unmarshal(file);
		// Gets the list of stock details.
		List<StockDetails> list = stockMarket.getStockDetails();
		ListSupport<StockDetails> stockDetails = new ListSupport<StockDetails>(StockDetails.class);
		stockDetails.addRange(list);
		// Creates an instance of MailMergeDataTable by specifying mail merge group name and IEnumerable collection.
		MailMergeDataTable dataTable = new MailMergeDataTable("StockDetails", stockDetails);
		return dataTable;
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
