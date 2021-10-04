package employeereport;

import java.util.List;
import java.io.*;
import javax.xml.bind.*;
import com.syncfusion.docio.*;
import com.syncfusion.javahelper.system.ConvertSupport;
import com.syncfusion.javahelper.system.collections.generic.ListSupport;

public class EmployeeReport {
	public static void main(String[] args) throws Exception {
		// Creating a new document.
		WordDocument document = new WordDocument();
		// Load template.
		document.open(getDataDir("EmployeesReportDemo.docx"), FormatType.Docx);
		// Create MailMergeDataTable.
		document.getMailMerge().MergeImageField.add("mergeField_EmployeeImage", new MergeImageFieldEventHandler() {
			ListSupport<MergeImageFieldEventHandler> delegateList = new ListSupport<MergeImageFieldEventHandler>(
					MergeImageFieldEventHandler.class);

			public void invoke(Object sender, MergeImageFieldEventArgs args) throws Exception {
				mergeField_EmployeeImage(sender, args);
			}

			public void dynamicInvoke(Object... args) throws Exception {
				mergeField_EmployeeImage((Object) args[0], (MergeImageFieldEventArgs) args[1]);
			}

			public void add(MergeImageFieldEventHandler delegate) throws Exception {
				if (delegate != null)
					delegateList.add(delegate);
			}

			public void remove(MergeImageFieldEventHandler delegate) throws Exception {
				if (delegate != null)
					delegateList.remove(delegate);
			}
		});
		// Execute Mail Merge with groups.
		MailMergeDataTable mailMergeDataTable = getMailMergeDataTableEmployee();
		document.getMailMerge().executeGroup(mailMergeDataTable);
		// Save and close the document.
		document.save("Employee Report.docx", FormatType.Docx);
		document.close();
		System.out.println("Word document generated successfully.");
	}

	/**
	 * Gets the image from disk during Merge.
	 * 
	 * @param sender Represents the object.
	 * @param args   Represents MergeImageField EventArguments
	 * 
	 */
	private static void mergeField_EmployeeImage(Object sender, MergeImageFieldEventArgs args) throws Exception {
		if ((args.getFieldName()).equals("Photo")) {
			String ProductFileName = args.getFieldValue().toString();
			byte[] bytes = ConvertSupport.fromBase64String(ProductFileName);
			ByteArrayInputStream stream = new ByteArrayInputStream(bytes);
			args.setImageStream(stream);
		}
	}

	/**
	 * 
	 * Gets the mail merge data table.
	 * 
	 */
	private static MailMergeDataTable getMailMergeDataTableEmployee() throws Exception {
		// Loads the XML file.
		File file = new File(getDataDir("EmployeesList.xml"));
		// Create a new instance for the JAXBContext.
		JAXBContext jaxbContext = JAXBContext.newInstance(Employees.class);
		// Reads the XML file.
		Unmarshaller jaxbUnmarshaller = jaxbContext.createUnmarshaller();
		Employees employees = (Employees) jaxbUnmarshaller.unmarshal(file);
		// Gets the list of employee details.
		List<Employee> employee = employees.getEmployees();
		ListSupport<Employee> employeeList = new ListSupport<Employee>(Employee.class);
		employeeList.addRange(employee);
		// Creates an instance of MailMergeDataTable by specifying mail merge group name and IEnumerable collection.	
		MailMergeDataTable dataTable = new MailMergeDataTable("Employees", employeeList);
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
