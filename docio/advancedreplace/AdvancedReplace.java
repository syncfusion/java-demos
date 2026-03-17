package advancedreplace;

import java.io.File;
import java.util.regex.Pattern;

import com.syncfusion.docio.*;
import com.syncfusion.javahelper.system.StringSupport;
import com.syncfusion.javahelper.system.text.regularExpressions.MatchSupport;

public class AdvancedReplace {
	public static void main(String[] args) throws Exception {

               // Creates a new Word document.
		WordDocument docSource1 = new WordDocument();
		WordDocument docSource2 = new WordDocument();
		WordDocument docMaster = new WordDocument();
		// Loads the template Word document.
		docSource1.open(getDataDir("SourceTemplate1.docx"),FormatType.Docx);
		
		docSource2.open(getDataDir("SourceTemplate2.docx"),FormatType.Docx);
		
		docMaster.open(getDataDir("MasterTemplate.docx"),FormatType.Docx);
		// Search for a string and store in TextSelection
                // The TextSelection copies a text segment with formatting.
		TextSelection selection1 = docSource1.find("PlaceHolder text is replaced with this formatted animated text",false,false);
		// Get the text segment to replace the text across multiple paragraphs.
		TextBodyPart replacePart = new TextBodyPart(docSource2);
		for(Object bodyItem_tempObj : docSource2.getLastSection().getBody().getChildEntities())
		{
		TextBodyItem bodyItem = (TextBodyItem)bodyItem_tempObj;
		replacePart.getBodyItems().add(bodyItem.clone());
		}
		// Replacing the placeholder inside Master Template with matches found while
                // search the two template documents. 
		docMaster.replace("PlaceHolder1",selection1,true,true,true);
		docMaster.replaceSingleLine((Pattern.compile(MatchSupport.trimPattern("PlaceHolder2Start:Suppliers/Vendors of Northwind." + "Customers of Northwind.Employee details of Northwind traders.The product information.The inventory details.The shippers." + "PO transactions i.e Purchase Order transactions.Sales Order transaction.Inventory transactions.Invoices.PlaceHolder2End"))),replacePart);
		// Saves and closes the Word document.
		docMaster.save("AdvancedReplace.docx", FormatType.Docx);
		docMaster.close();
				
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
