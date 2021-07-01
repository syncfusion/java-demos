package nestedmailmerge;

import java.io.File;

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
		// Gets the employee details as “IEnumerable” collection.
		ListSupport<EmployeeDetails> employees = new ListSupport<EmployeeDetails>(EmployeeDetails.class);
		// Gets the customer details as “IEnumerable” collection.
		ListSupport<CustomerDetails> customers = new ListSupport<CustomerDetails>(CustomerDetails.class);
		// Gets the order details as “IEnumerable” collection.
		ListSupport<OrderDetails> orders = new ListSupport<OrderDetails>(OrderDetails.class);
		FileStreamSupport stream = new FileStreamSupport(getDataDir("Employees.xml"), FileMode.OpenOrCreate);
		XmlReaderSupport reader = XmlReaderSupport.create(stream);
		if (reader == null)
			throw new Exception("reader");
		while (reader.getNodeType().getEnumValue() != XmlNodeType.Element.getEnumValue())
			reader.read();
		if (!(reader.getLocalName() == "EmployeesList"))
			throw new Exception("Unexpected xml tag " + reader.getLocalName());
		reader.read();
		while (reader.getNodeType().getEnumValue() == XmlNodeType.Whitespace.getEnumValue())
			reader.read();
		while (!(reader.getLocalName() == "EmployeesList")) {
			if (reader.getNodeType().getEnumValue() == XmlNodeType.Element.getEnumValue()) {
				switch ((reader.getLocalName()) == null ? "string_null_value" : (reader.getLocalName())) {
				case "Employees":

					employees.add(getEmployee(reader, customers, orders));
					break;
				}
			} else

			{
				reader.read();
				if ((reader.getLocalName() == "EmployeesList")
						&& reader.getNodeType().getEnumValue() == XmlNodeType.EndElement.getEnumValue())
					break;
			}
		}
		reader.close();
		stream.close();
		// Creates an instance of the MailMergeDataSet.
		MailMergeDataSet dataSet = new MailMergeDataSet();
		dataSet.add(new MailMergeDataTable("Employees", employees));
		dataSet.add(new MailMergeDataTable("Customers", customers));
		dataSet.add(new MailMergeDataTable("Orders", orders));
		return dataSet;
	}

	/**
	 * 
	 * Gets the employee.
	 * 
	 * @param reader    The reader.
	 * @param customers The customers.
	 * @param orders    The orders.
	 * @return
	 * @exception System.Exception reader
	 * @exception XmlException     Unexpected xml tag + reader.LocalName
	 */
	private static EmployeeDetails getEmployee(XmlReaderSupport reader, ListSupport<CustomerDetails> customers,
			ListSupport<OrderDetails> orders) throws Exception {
		if (reader == null)
			throw new Exception("reader");
		while (reader.getNodeType().getEnumValue() != XmlNodeType.Element.getEnumValue())
			reader.read();
		if (!(reader.getLocalName() == "Employees"))
			throw new Exception("Unexpected xml tag " + reader.getLocalName());
		reader.read();
		while (reader.getNodeType().getEnumValue() == XmlNodeType.Whitespace.getEnumValue())
			reader.read();
		EmployeeDetails employee = new EmployeeDetails();
		while (!(reader.getLocalName() == "Employees")) {
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
				case "Address":

					employee.setAddress(reader.readContentAsString());
					break;
				case "City":

					employee.setCity(reader.readContentAsString());
					break;
				case "Country":

					employee.setCountry(reader.readContentAsString());
					break;
				case "Extension":

					employee.setExtension(reader.readContentAsString());
					break;
				case "Customers":

					customers.add(getCustomer(reader, orders));
					break;
				default:

					reader.skip();
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
		return employee;
	}

	/**
	 * 
	 * Gets the customer.
	 * 
	 * @param reader The reader.
	 * @param orders The orders.
	 * @return
	 * @exception System.Exception reader
	 * @exception XmlException     Unexpected xml tag + reader.LocalName
	 */
	private static CustomerDetails getCustomer(XmlReaderSupport reader, ListSupport<OrderDetails> orders)
			throws Exception {
		if (reader == null)
			throw new Exception("reader");
		while (reader.getNodeType().getEnumValue() != XmlNodeType.Element.getEnumValue())
			reader.read();
		if (!(reader.getLocalName() == "Customers"))
			throw new Exception("Unexpected xml tag " + reader.getLocalName());
		reader.read();
		while (reader.getNodeType().getEnumValue() == XmlNodeType.Whitespace.getEnumValue())
			reader.read();
		CustomerDetails customer = new CustomerDetails();
		while (!(reader.getLocalName() == "Customers")) {
			if (reader.getNodeType().getEnumValue() == XmlNodeType.Element.getEnumValue()) {
				switch ((reader.getLocalName()) == null ? "string_null_value" : (reader.getLocalName())) {
				case "EmployeeID":

					customer.setEmployeeID(reader.readContentAsString());
					break;
				case "CustomerID":

					customer.setCustomerID(reader.readContentAsString());
					break;
				case "CompanyName":

					customer.setCompanyName(reader.readContentAsString());
					break;
				case "ContactName":

					customer.setContactName(reader.readContentAsString());
					break;
				case "City":

					customer.setCity(reader.readContentAsString());
					break;
				case "Country":

					customer.setCountry(reader.readContentAsString());
					break;
				case "Orders":

					orders.add(getOrders(reader));
					break;
				}
				reader.read();
			} else

			{
				reader.read();
				if ((reader.getLocalName() == "Customers")
						&& reader.getNodeType().getEnumValue() == XmlNodeType.EndElement.getEnumValue())
					break;
			}
		}
		return customer;
	}

	/**
	 * 
	 * Gets the mail merge data table.
	 * 
	 * @return
	 * @exception System.Exception reader
	 * @exception XmlException     Unexpected xml tag + reader.LocalName
	 */
	private MailMergeDataTable getMailMergeDataTable(String basePath) throws Exception {
		// Gets the employee details implicit as “IEnumerable” collection.
		ListSupport<EmployeeDetailsImplicit> employees = new ListSupport<EmployeeDetailsImplicit>(
				EmployeeDetailsImplicit.class);
		FileStreamSupport stream = new FileStreamSupport(basePath + "/DocIO/Employees.xml", FileMode.OpenOrCreate);
		XmlReaderSupport reader = XmlReaderSupport.create(stream);
		if (reader == null)
			throw new Exception("reader");
		while (reader.getNodeType().getEnumValue() != XmlNodeType.Element.getEnumValue())
			reader.read();
		if (!(reader.getLocalName() == "EmployeesList"))
			throw new Exception("Unexpected xml tag " + reader.getLocalName());
		reader.read();
		while (reader.getNodeType().getEnumValue() == XmlNodeType.Whitespace.getEnumValue())
			reader.read();
		while (!(reader.getLocalName() == "EmployeesList")) {
			if (reader.getNodeType().getEnumValue() == XmlNodeType.Element.getEnumValue()) {
				switch ((reader.getLocalName()) == null ? "string_null_value" : (reader.getLocalName())) {
				case "Employees":

					employees.add(getEmployee(reader));
					break;
				}
			} else

			{
				reader.read();
				if ((reader.getLocalName() == "EmployeesList")
						&& reader.getNodeType().getEnumValue() == XmlNodeType.EndElement.getEnumValue())
					break;
			}
		}
		reader.close();
		stream.close();
		// Creates an instance of “MailMergeDataTable” by specifying mail merge group name and “IEnumerable” collection.
		MailMergeDataTable dataTable = new MailMergeDataTable("Employees", employees);
		return dataTable;
	}

	/**
	 * 
	 * Gets the employee.
	 * 
	 * @param reader The reader.
	 * @return
	 * @exception System.Exception reader
	 * @exception XmlException     Unexpected xml tag + reader.LocalName
	 */
	private EmployeeDetailsImplicit getEmployee(XmlReaderSupport reader) throws Exception {
		if (reader == null)
			throw new Exception("reader");
		while (reader.getNodeType().getEnumValue() != XmlNodeType.Element.getEnumValue())
			reader.read();
		if (!(reader.getLocalName() == "Employees"))
			throw new Exception("Unexpected xml tag " + reader.getLocalName());
		reader.read();
		while (reader.getNodeType().getEnumValue() == XmlNodeType.Whitespace.getEnumValue())
			reader.read();
		EmployeeDetailsImplicit employee = new EmployeeDetailsImplicit();
		while (!(reader.getLocalName() == "Employees")) {
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
				case "Address":

					employee.setAddress(reader.readContentAsString());
					break;
				case "City":

					employee.setCity(reader.readContentAsString());
					break;
				case "Country":

					employee.setCountry(reader.readContentAsString());
					break;
				case "Extension":

					employee.setExtension(reader.readContentAsString());
					break;
				case "Customers":

					employee.getCustomers().add(getCustomer(reader));
					break;
				default:

					reader.skip();
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
		return employee;
	}

	/**
	 * 
	 * Gets the customer.
	 * 
	 * @param reader The reader.
	 * @return
	 * @exception System.Exception reader
	 * @exception XmlException     Unexpected xml tag + reader.LocalName
	 */
	private CustomerDetailsImplicit getCustomer(XmlReaderSupport reader) throws Exception {
		if (reader == null)
			throw new Exception("reader");
		while (reader.getNodeType().getEnumValue() != XmlNodeType.Element.getEnumValue())
			reader.read();
		if (!(reader.getLocalName() == "Customers"))
			throw new Exception("Unexpected xml tag " + reader.getLocalName());
		reader.read();
		while (reader.getNodeType().getEnumValue() == XmlNodeType.Whitespace.getEnumValue())
			reader.read();
		CustomerDetailsImplicit customer = new CustomerDetailsImplicit();
		while (!(reader.getLocalName() == "Customers")) {
			if (reader.getNodeType().getEnumValue() == XmlNodeType.Element.getEnumValue()) {
				switch ((reader.getLocalName()) == null ? "string_null_value" : (reader.getLocalName())) {
				case "EmployeeID":

					customer.setEmployeeID(reader.readContentAsString());
					break;
				case "CustomerID":

					customer.setCustomerID(reader.readContentAsString());
					break;
				case "CompanyName":

					customer.setCompanyName(reader.readContentAsString());
					break;
				case "ContactName":

					customer.setContactName(reader.readContentAsString());
					break;
				case "City":

					customer.setCity(reader.readContentAsString());
					break;
				case "Country":

					customer.setCountry(reader.readContentAsString());
					break;
				case "Orders":

					customer.getOrders().add(getOrders(reader));
					break;
				}
				reader.read();
			} else

			{
				reader.read();
				if ((reader.getLocalName() == "Customers")
						&& reader.getNodeType().getEnumValue() == XmlNodeType.EndElement.getEnumValue())
					break;
			}
		}
		return customer;
	}

	/**
	 * 
	 * Gets the orders.
	 * 
	 * @param reader The reader.
	 * @return
	 * @exception System.Exception reader
	 * @exception XmlException     Unexpected xml tag + reader.LocalName
	 */
	private static OrderDetails getOrders(XmlReaderSupport reader) throws Exception {
		if (reader == null)
			throw new Exception("reader");
		while (reader.getNodeType().getEnumValue() != XmlNodeType.Element.getEnumValue())
			reader.read();
		if (!(reader.getLocalName() == "Orders"))
			throw new Exception("Unexpected xml tag " + reader.getLocalName());
		reader.read();
		while (reader.getNodeType().getEnumValue() == XmlNodeType.Whitespace.getEnumValue())
			reader.read();
		OrderDetails order = new OrderDetails();
		while (!(reader.getLocalName() == "Orders")) {
			if (reader.getNodeType().getEnumValue() != XmlNodeType.EndElement.getEnumValue()) {
				switch ((reader.getLocalName()) == null ? "string_null_value" : (reader.getLocalName())) {
				case "OrderID":

					order.setOrderID(reader.readContentAsString());
					break;
				case "CustomerID":

					order.setCustomerID(reader.readContentAsString());
					break;
				case "OrderDate":

					order.setOrderDate(reader.readContentAsString());
					break;
				case "RequiredDate":

					order.setRequiredDate(reader.readContentAsString());
					break;
				case "ShippedDate":

					order.setShippedDate(reader.readContentAsString());
					break;
				}
				reader.read();
			} else

			{
				reader.read();
				if ((reader.getLocalName() == "Orders")
						&& reader.getNodeType().getEnumValue() == XmlNodeType.EndElement.getEnumValue())
					break;
			}
		}
		return order;
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
