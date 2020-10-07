package employeereport;

import java.io.*;
import com.syncfusion.docio.*;
import com.syncfusion.javahelper.system.*;
import com.syncfusion.javahelper.system.collections.generic.ListSupport;
import com.syncfusion.javahelper.system.io.*;
import com.syncfusion.javahelper.system.xml.*;

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
		ListSupport<Employees> employees = new ListSupport<Employees>(Employees.class);
		FileStreamSupport fs = new FileStreamSupport(getDataDir("EmployeesList.xml"), FileMode.Open,
				FileAccess.Read);
		XmlReaderSupport reader = XmlReaderSupport.create(fs);
		if (reader == null)
			throw new Exception("reader");
		while (reader.getNodeType().getEnumValue() != XmlNodeType.Element.getEnumValue())
			reader.read();
		if (!(reader.getLocalName() == "Employees"))
			throw new Exception("Unexpected xml tag " + reader.getLocalName());
		reader.read();
		while (reader.getNodeType().getEnumValue() == XmlNodeType.Whitespace.getEnumValue())
			reader.read();
		while (!(reader.getLocalName() == "Employees")) {
			if (reader.getNodeType().getEnumValue() == XmlNodeType.Element.getEnumValue()) {
				switch ((reader.getLocalName()) == null ? "string_null_value" : (reader.getLocalName())) {
				case "Employee":

					employees.add(getEmployees(reader));
					break;
				}
			} else

			{
				reader.read();
				if ((reader.getLocalName() == "Employees")
						&& reader.getNodeType().getEnumValue() == XmlNodeType.EndElement.getEnumValue())
					break;
			}
		}
		MailMergeDataTable dataTable = new MailMergeDataTable("Employees", employees);
		reader.close();
		fs.close();
		return dataTable;
	}

	/**
	 * 
	 * Gets the employees.
	 * 
	 * @param reader The reader.
	 */
	private static Employees getEmployees(XmlReaderSupport reader) throws Exception {
		if (reader == null)
			throw new Exception("reader");
		while (reader.getNodeType().getEnumValue() != XmlNodeType.Element.getEnumValue())
			reader.read();
		if (!(reader.getLocalName() == "Employee"))
			throw new Exception("Unexpected xml tag " + reader.getLocalName());
		reader.read();
		while (reader.getNodeType().getEnumValue() == XmlNodeType.Whitespace.getEnumValue())
			reader.read();
		Employees employee = new Employees();
		while (!(reader.getLocalName() == "Employee")) {
			if (reader.getNodeType().getEnumValue() == XmlNodeType.Element.getEnumValue()) {
				switch ((reader.getLocalName()) == null ? "string_null_value" : (reader.getLocalName())) {
				case "EmployeeID":

					employee.setEmployeeID(reader.readContentAsString());
					break;
				case "LastName":

					employee.setLastName(reader.readContentAsString());
					break;
				case "FirstName":

					employee.setFirstName(reader.readContentAsString());
					break;
				case "Title":

					employee.setTitle(reader.readContentAsString());
					break;
				case "TitleOfCourtesy":

					employee.setTitleOfCourtesy(reader.readContentAsString());
					break;
				case "BirthDate":

					employee.setBirthDate(reader.readContentAsString());
					break;
				case "HireDate":

					employee.setHireDate(reader.readContentAsString());
					break;
				case "Address":

					employee.setAddress(reader.readContentAsString());
					break;
				case "City":

					employee.setCity(reader.readContentAsString());
					break;
				case "Region":

					employee.setRegion(reader.readContentAsString());
					break;
				case "PostalCode":

					employee.setPostalCode(reader.readContentAsString());
					break;
				case "Country":

					employee.setCountry(reader.readContentAsString());
					break;
				case "HomePhone":

					employee.setHomePhone(reader.readContentAsString());
					break;
				case "Extension":

					employee.setExtension(reader.readContentAsString());
					break;
				case "Photo":

					employee.setPhoto(reader.readContentAsString());
					break;
				case "Notes":

					employee.setNotes(reader.readContentAsString());
					break;
				case "ReportsTo":

					employee.setReportsTo(reader.readContentAsString());
					break;
				default:

					reader.skip();
					break;
				}
			} else

			{
				reader.read();
				if ((reader.getLocalName() == "Employee")
						&& reader.getNodeType().getEnumValue() == XmlNodeType.EndElement.getEnumValue())
					break;
			}
		}
		return employee;
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
