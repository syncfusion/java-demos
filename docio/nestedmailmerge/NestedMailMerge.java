package nestedmailmerge;

import java.io.File;
import java.util.List;

import javax.xml.bind.JAXBContext;
import javax.xml.bind.Unmarshaller;

import com.syncfusion.docio.*;
import com.syncfusion.javahelper.system.*;
import com.syncfusion.javahelper.system.collections.DictionaryEntrySupport;
import com.syncfusion.javahelper.system.collections.generic.ListSupport;
import com.syncfusion.javahelper.system.io.*;
import com.syncfusion.javahelper.system.xml.*;

public class NestedMailMerge {
	public static void main(String[] args) throws Exception {
		// Creates a new document.
		WordDocument document = new WordDocument();
		// Loads template
		String basePath = getDataDir("Template_Report.docx");
		document.open(basePath, FormatType.Docx);
		// Retrieves the mail merge data.
		MailMergeDataSet dataSet = getMailMergeDataSet();
		ListSupport<DictionaryEntrySupport> commands = new ListSupport<DictionaryEntrySupport>(
				DictionaryEntrySupport.class);
		// DictionaryEntry contain "Source table" (KEY) and "Command" (VALUE)
		DictionaryEntrySupport entry = new DictionaryEntrySupport("Employees", "");
		commands.add(entry);
		// Retrieves the customer details.
		entry = new DictionaryEntrySupport("Customers", "EmployeeID = %Employees.EmployeeID%");
		commands.add(entry);
		// Retrieves the order details.
		entry = new DictionaryEntrySupport("Orders", "CustomerID = %Customers.CustomerID%");
		commands.add(entry);
		// Executes the mail merge as nested group.
		document.getMailMerge().executeNestedGroup(dataSet, commands);
		// Save and close the document.
		document.save("Nested Mail Merge.docx", FormatType.Docx);
		document.close();
		System.out.println("Word document generated successfully.");
	}

	/**
	 * 
	 * Gets the mail merge data set.
	 * 
	 * @return
	 * @exception System.Exception reader
	 * @exception XmlException     Unexpected xml tag + reader.LocalName
	 */
	private static MailMergeDataSet getMailMergeDataSet() throws Exception {
		// Loads the XML file.
		File file = new File(getDataDir("Employees.xml"));
		// Create a new instance for the JAXBContext.
		JAXBContext jaxbContext = JAXBContext.newInstance(Employees.class);
		// Reads the XML file.
		Unmarshaller jaxbUnmarshaller = jaxbContext.createUnmarshaller();
		Employees employees = (Employees) jaxbUnmarshaller.unmarshal(file);
		// Gets the employee details as “IEnumerable” collection.
		ListSupport<EmployeeDetails> employeeList = new ListSupport<EmployeeDetails>(EmployeeDetails.class);
		// Gets the list of employee details.
		List<EmployeeDetails> list = employees.getEmployees();
		employeeList.addRange(list);
		// Gets the customer details as IEnumerable collection.
		ListSupport<CustomerDetails> customers = employees.getCustomers();
		// Gets the order details as IEnumerable collection.
		ListSupport<OrderDetails> orders = employees.getOrders();
		// Creates an instance of the MailMergeDataSet.
		MailMergeDataSet dataSet = new MailMergeDataSet();
		dataSet.add(new MailMergeDataTable("Employees", employeeList));
		dataSet.add(new MailMergeDataTable("Customers", customers));
		dataSet.add(new MailMergeDataTable("Orders", orders));
		return dataSet;
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
