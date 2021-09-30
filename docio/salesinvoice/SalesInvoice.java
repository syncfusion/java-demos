package salesinvoice;

import java.io.File;
import java.util.*;
import com.syncfusion.docio.*;
import com.syncfusion.javahelper.system.*;
import com.syncfusion.javahelper.system.collections.generic.ListSupport;
import com.syncfusion.javahelper.system.drawing.*;
import com.syncfusion.javahelper.system.io.*;
import com.syncfusion.javahelper.system.xml.*;

public class SalesInvoice {
	public static void main(String[] args) throws Exception {
		// Loads the template.
		WordDocument doc = new WordDocument(getDataDir("SalesInvoiceDemo.docx"), FormatType.Docx);
		// Creates MailMergeDataTable
		MailMergeDataTable mailMergeDataTableOrder = getTestOrder(10248);
		MailMergeDataTable mailMergeDataTableOrderTotals = getTestOrderTotals(10248);
		MailMergeDataTable mailMergeDataTableOrderDetails = getTestOrderDetails(10248);
		// Executes Mail Merge with groups.
		doc.getMailMerge().executeGroup(mailMergeDataTableOrder);
		doc.getMailMerge().executeGroup(mailMergeDataTableOrderTotals);
		// Using Merge events to do conditional formatting during runtime.
		doc.getMailMerge().MergeField.add("mailMerge_MergeField", new MergeFieldEventHandler() {
			ListSupport<MergeFieldEventHandler> delegateList = new ListSupport<MergeFieldEventHandler>(
					MergeFieldEventHandler.class);
			// Represents event handling for MergeFieldEventHandlerCollection.
			public void invoke(Object sender, MergeFieldEventArgs args) throws Exception {
				mailMerge_MergeField(sender, args);
			}
            // Represents the method that handles MergeField event.
			public void dynamicInvoke(Object... args) throws Exception {
				mailMerge_MergeField((Object) args[0], (MergeFieldEventArgs) args[1]);
			}
			// Represents the method that handles MergeField event to add collection item.
			public void add(MergeFieldEventHandler delegate) throws Exception {
				if (delegate != null)
					delegateList.add(delegate);
			}
			// Represents the method that handles MergeField event to remove collection item.
			public void remove(MergeFieldEventHandler delegate) throws Exception {
				if (delegate != null)
					delegateList.remove(delegate);
			}
		});
		// Executes Mail Merge with groups.
		doc.getMailMerge().executeGroup(mailMergeDataTableOrderDetails);
		// Save and close the document.
		doc.save("Sales Invoice.docx", FormatType.Docx);
		doc.close();
		System.out.println("Word document generated successfully.");
	}

	/**
	 * Conditionally format data during Merge.
	 * 
	 * @param sender Represents the object.
	 * @param args   Represents MergeField EventArguments
	 */
	private static void mailMerge_MergeField(Object sender, MergeFieldEventArgs args) throws Exception {
		if (Integer.compare((args.getRowIndex() % 2),0)==0) {
			// Sets text color to the merge field event.
			args.getCharacterFormat().setTextColor(ColorSupport.getDarkBlue());
		}
	}

	/**
	 * Get Test Order details
	 * 
	 * @param TestOrderId Represents test order id
	 * 
	 */
	private static MailMergeDataTable getTestOrder(int TestOrderId) throws Exception {
		// Gets the test order as “IEnumerable” collection.
		ListSupport<TestOrder> orders = new ListSupport<TestOrder>(TestOrder.class);
		FileStreamSupport fs = new FileStreamSupport(getDataDir("TestOrder.xml"), FileMode.Open, FileAccess.Read);
		XmlReaderSupport reader = XmlReaderSupport.create(fs);
		if (reader == null)
			throw new Exception("reader");
		while (reader.getNodeType().getEnumValue() != XmlNodeType.Element.getEnumValue())
			reader.read();
		if (!(reader.getLocalName().equals("TestOrders")))
			throw new Exception("Unexpected xml tag " + reader.getLocalName());
		reader.read();
		while (reader.getNodeType().getEnumValue() == XmlNodeType.Whitespace.getEnumValue())
			reader.read();
		while (!(reader.getLocalName().equals("TestOrders"))) {
			if (reader.getNodeType().getEnumValue() == XmlNodeType.Element.getEnumValue()) {
				switch ((reader.getLocalName()) == null ? "string_null_value" : (reader.getLocalName())) {
				case "TestOrder":
                    // Gets the test order.
					TestOrder testOrder = getTestOrder(reader);
					if (testOrder.getOrderID().equals(Int32Support.toString(TestOrderId))) {
						orders.add(testOrder);
						break;
					}
					break;
				}
			} else

			{
				reader.read();
				if ((reader.getLocalName().equals("TestOrders"))
						&& reader.getNodeType().getEnumValue() == XmlNodeType.EndElement.getEnumValue())
					break;
			}
		}
		// Creates an instance of “MailMergeDataTable” by specifying mail merge group name and “IEnumerable” collection.
		MailMergeDataTable dataTable = new MailMergeDataTable("Orders", orders);
		reader.close();
		fs.close();
		return dataTable;
	}

	/**
	 * Gets test order details
	 * 
	 * @param TestOrderId Represents test order id
	 * 
	 */
	private static MailMergeDataTable getTestOrderDetails(int TestOrderId) throws Exception {
		// Gets the test order details as “IEnumerable” collection.
		ListSupport<TestOrderDetail> orders = new ListSupport<TestOrderDetail>(TestOrderDetail.class);
		FileStreamSupport fs = new FileStreamSupport(getDataDir("TestOrderDetails.xml"), FileMode.Open,
				FileAccess.Read);
		XmlReaderSupport reader = XmlReaderSupport.create(fs);
		if (reader == null)
			throw new Exception("reader");
		while (reader.getNodeType().getEnumValue() != XmlNodeType.Element.getEnumValue())
			reader.read();
		if (!(reader.getLocalName().equals("TestOrderDetails")))
			throw new Exception("Unexpected xml tag " + reader.getLocalName());
		reader.read();
		while (reader.getNodeType().getEnumValue() == XmlNodeType.Whitespace.getEnumValue())
			reader.read();
		while (!(reader.getLocalName().equals("TestOrderDetails"))) {
			if (reader.getNodeType().getEnumValue() == XmlNodeType.Element.getEnumValue()) {
				switch ((reader.getLocalName()) == null ? "string_null_value" : (reader.getLocalName())) {
				case "TestOrderDetail":
					// Gets the test order details.
					TestOrderDetail testOrder = getTestOrderDetail(reader);
					if ((testOrder.getOrderID().equals(Int32Support.toString(TestOrderId)))) {
						orders.add(testOrder);
						break;
					}
					break;
				}
			} else

			{
				reader.read();
				if ((reader.getLocalName().equals("TestOrderDetails"))
						&& reader.getNodeType().getEnumValue() == XmlNodeType.EndElement.getEnumValue())
					break;
			}
		}
		// Creates an instance of “MailMergeDataTable” by specifying mail merge group name and “IEnumerable” collection.
		MailMergeDataTable dataTable = new MailMergeDataTable("Order", orders);
		reader.close();
		fs.close();
		return dataTable;
	}

	/**
	 * Gets test order totals
	 * 
	 * @param TestOrderId Represents test order id
	 * 
	 */
	private static MailMergeDataTable getTestOrderTotals(int TestOrderId) throws Exception {
		// Gets the test order total details as “IEnumerable” collection.
		ListSupport<TestOrderTotal> orders = new ListSupport<TestOrderTotal>(TestOrderTotal.class);
		FileStreamSupport fs = new FileStreamSupport(getDataDir("OrderTotals.xml"), FileMode.Open,
				FileAccess.Read);
		XmlReaderSupport reader = XmlReaderSupport.create(fs);
		if (reader == null)
			throw new Exception("reader");
		while (reader.getNodeType().getEnumValue() != XmlNodeType.Element.getEnumValue())
			reader.read();
		if (!(reader.getLocalName().equals("TestOrderTotals")))
			throw new Exception("Unexpected xml tag " + reader.getLocalName());
		reader.read();
		while (reader.getNodeType().getEnumValue() == XmlNodeType.Whitespace.getEnumValue())
			reader.read();
		while (!(reader.getLocalName().equals("TestOrderTotals"))) {
			if (reader.getNodeType().getEnumValue() == XmlNodeType.Element.getEnumValue()) {
				switch ((reader.getLocalName()) == null ? "string_null_value" : (reader.getLocalName())) {
				case "TestOrderTotal":
                    // Gets the test order total.
					TestOrderTotal testOrder = getTestOrderTotal(reader);
					if ((testOrder.getOrderID().equals(Int32Support.toString(TestOrderId)))) {
						orders.add(testOrder);
						break;
					}
					break;
				}
				reader.read();
			} else

			{
				reader.read();
				if ((reader.getLocalName().equals("TestOrderTotals"))
						&& reader.getNodeType().getEnumValue() == XmlNodeType.EndElement.getEnumValue())
					break;
			}
		}
		// Creates an instance of “MailMergeDataTable” by specifying mail merge group name and “IEnumerable” collection.
		MailMergeDataTable dataTable = new MailMergeDataTable("OrderTotals", orders);
		reader.close();
		fs.close();
		return dataTable;
	}

	/**
	 * 
	 * Gets the test order.
	 * 
	 * @param reader The reader.
	 */
	private static TestOrder getTestOrder(XmlReaderSupport reader) throws Exception {
		if (reader == null)
			throw new Exception("reader");
		while (reader.getNodeType().getEnumValue() != XmlNodeType.Element.getEnumValue())
			reader.read();
		if (!(reader.getLocalName().equals("TestOrder")))
			throw new Exception("Unexpected xml tag " + reader.getLocalName());
		reader.read();
		while (reader.getNodeType().getEnumValue() == XmlNodeType.Whitespace.getEnumValue())
			reader.read();
		TestOrder testOrder = new TestOrder();
		while (!(reader.getLocalName().equals("TestOrder"))) {
			if (reader.getNodeType().getEnumValue() == XmlNodeType.Element.getEnumValue()) {
				switch ((reader.getLocalName()) == null ? "string_null_value" : (reader.getLocalName())) {
				case "ShipName":

					testOrder.setShipName(reader.readContentAsString());
					break;
				case "ShipAddress":

					testOrder.setShipAddress(reader.readContentAsString());
					break;
				case "ShipCity":

					testOrder.setShipCity(reader.readContentAsString());
					break;
				case "ShipPostalCode":

					testOrder.setShipPostalCode(reader.readContentAsString());
					break;
				case "ShipCountry":

					testOrder.setShipCountry(reader.readContentAsString());
					break;
				case "PostalCode":

					testOrder.setPostalCode(reader.readContentAsString());
					break;
				case "CustomerID":

					testOrder.setCustomerID(reader.readContentAsString());
					break;
				case "Customers.CompanyName":

					testOrder.setCustomers_CompanyName(reader.readContentAsString());
					break;
				case "HomePage":

					testOrder.setSalesperson(reader.readContentAsString());
					break;
				case "Address":

					testOrder.setAddress(reader.readContentAsString());
					break;
				case "City":

					testOrder.setCity(reader.readContentAsString());
					break;
				case "Country":

					testOrder.setCountry(reader.readContentAsString());
					break;
				case "OrderID":

					testOrder.setOrderID(reader.readContentAsString());
					break;
				case "OrderDate":

					testOrder.setOrderDate(reader.readContentAsString());
					break;
				case "RequiredDate":

					testOrder.setRequiredDate(reader.readContentAsString());
					break;
				case "ShippedDate":

					testOrder.setShippedDate(reader.readContentAsString());
					break;
				case "Shippers.CompanyName":

					testOrder.setShippers_CompanyName(reader.readContentAsString());
					break;
				default:

					reader.skip();
					break;
				}
			} else

			{
				reader.read();
				if ((reader.getLocalName().equals("TestOrders"))
						&& reader.getNodeType().getEnumValue() == XmlNodeType.EndElement.getEnumValue())
					break;
			}
		}
		return testOrder;
	}

	/**
	 * Gets TestOrder ID
	 *
	 */
	@SuppressWarnings("unchecked")
	private ArrayList getTestOrderID() throws Exception {
		FileStreamSupport fs = new FileStreamSupport(getDataDir("TestOrder.xml"), FileMode.Open, FileAccess.Read);
		XmlReaderSupport reader = XmlReaderSupport.create(fs);
		if (reader == null)
			throw new Exception("reader");
		while (reader.getNodeType().getEnumValue() != XmlNodeType.Element.getEnumValue())
			reader.read();
		if (!(reader.getLocalName().equals("TestOrders")))
			throw new Exception("Unexpected xml tag " + reader.getLocalName());
		reader.read();
		while (reader.getNodeType().getEnumValue() == XmlNodeType.Whitespace.getEnumValue())
			reader.read();
		ArrayList orderId = new ArrayList();
		while (!(reader.getLocalName().equals("TestOrders"))) {
			if (reader.getNodeType().getEnumValue() == XmlNodeType.Element.getEnumValue()) {
				switch ((reader.getLocalName()) == null ? "string_null_value" : (reader.getLocalName())) {
				case "OrderID":

					orderId.add(reader.readContentAsString());
					break;
				default:

					break;
				}
				reader.read();
			} else

			{
				reader.read();
				if ((reader.getLocalName().equals("TestOrders"))
						&& reader.getNodeType().getEnumValue() == XmlNodeType.EndElement.getEnumValue())
					break;
			}
		}
		return orderId;
	}

	/**
	 * 
	 * Gets the test order details.
	 * 
	 * @param reader The reader.
	 */
	private static TestOrderDetail getTestOrderDetail(XmlReaderSupport reader) throws Exception {
		if (reader == null)
			throw new Exception("reader");
		while (reader.getNodeType().getEnumValue() != XmlNodeType.Element.getEnumValue())
			reader.read();
		if (!(reader.getLocalName().equals("TestOrderDetail")))
			throw new Exception("Unexpected xml tag " + reader.getLocalName());
		reader.read();
		while (reader.getNodeType().getEnumValue() == XmlNodeType.Whitespace.getEnumValue())
			reader.read();
		TestOrderDetail testOrderDetail = new TestOrderDetail();
		while (!(reader.getLocalName().equals("TestOrderDetail"))) {
			if (reader.getNodeType().getEnumValue() == XmlNodeType.Element.getEnumValue()) {
				switch ((reader.getLocalName()) == null ? "string_null_value" : (reader.getLocalName())) {
				case "OrderID":

					testOrderDetail.setOrderID(reader.readContentAsString());
					break;
				case "ProductID":

					testOrderDetail.setProductID(reader.readContentAsString());
					break;
				case "ProductName":

					testOrderDetail.setProductName(reader.readContentAsString());
					break;
				case "UnitPrice":

					testOrderDetail.setUnitPrice(reader.readContentAsString());
					break;
				case "Quantity":

					testOrderDetail.setQuantity(reader.readContentAsString());
					break;
				case "Discount":

					testOrderDetail.setDiscount(reader.readContentAsString());
					break;
				case "ExtendedPrice":

					testOrderDetail.setExtendedPrice(reader.readContentAsString());
					break;
				default:

					reader.skip();
					break;
				}
			} else

			{
				reader.read();
				if ((reader.getLocalName().equals("TestOrderDetail"))
						&& reader.getNodeType().getEnumValue() == XmlNodeType.EndElement.getEnumValue())
					break;
			}
		}
		return testOrderDetail;
	}

	/**
	 * 
	 * Gets the test order total.
	 * 
	 * @param reader The reader.
	 */
	private static TestOrderTotal getTestOrderTotal(XmlReaderSupport reader) throws Exception {
		if (reader == null)
			throw new Exception("reader");
		while (reader.getNodeType().getEnumValue() != XmlNodeType.Element.getEnumValue())
			reader.read();
		if (!(reader.getLocalName().equals("TestOrderTotal")))
			throw new Exception("Unexpected xml tag " + reader.getLocalName());
		reader.read();
		while (reader.getNodeType().getEnumValue() == XmlNodeType.Whitespace.getEnumValue())
			reader.read();
		TestOrderTotal testOrderTotal = new TestOrderTotal();
		while (!(reader.getLocalName().equals("TestOrderTotal"))) {
			if (reader.getNodeType().getEnumValue() == XmlNodeType.Element.getEnumValue()) {
				switch ((reader.getLocalName()) == null ? "string_null_value" : (reader.getLocalName())) {
				case "OrderID":

					testOrderTotal.setOrderID(reader.readContentAsString());
					break;
				case "Subtotal":

					testOrderTotal.setSubtotal(reader.readContentAsString());
					break;
				case "Freight":

					testOrderTotal.setFreight(reader.readContentAsString());
					break;
				case "Total":

					testOrderTotal.setTotal(reader.readContentAsString());
					break;
				default:

					reader.skip();
					break;
				}
			} else

			{
				reader.read();
				if ((reader.getLocalName().equals("TestOrderTotal"))
						&& reader.getNodeType().getEnumValue() == XmlNodeType.EndElement.getEnumValue())
					break;
			}
		}
		return testOrderTotal;
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
