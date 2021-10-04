package salesinvoice;

import java.io.File;
import java.util.*;

import javax.xml.bind.*;

import com.syncfusion.docio.*;
import com.syncfusion.javahelper.system.*;
import com.syncfusion.javahelper.system.collections.generic.ListSupport;
import com.syncfusion.javahelper.system.drawing.*;

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
		// Loads the XML file.
		File file = new File(getDataDir("TestOrder.xml"));
		// Create a new instance for the JAXBContext.
		JAXBContext jaxbContext = JAXBContext.newInstance(TestOrders.class);
		// Reads the XML file.
		Unmarshaller jaxbUnmarshaller = jaxbContext.createUnmarshaller();
		TestOrders testOrders = (TestOrders) jaxbUnmarshaller.unmarshal(file);
		// Gets the list of test order details.
		List<TestOrder> list = testOrders.getTestOrder();
		ListSupport<TestOrder> testOrderList = new ListSupport<TestOrder>(TestOrder.class);
		for (TestOrder testOrder : list) {
			if (testOrder.getOrderID().equals(Int32Support.toString(TestOrderId))) {				
				testOrderList.add(testOrder);
			}
		}
		// Creates an instance of MailMergeDataTable by specifying mail merge group name and IEnumerable collection.
		MailMergeDataTable dataTable = new MailMergeDataTable("Orders", testOrderList);
		return dataTable;
	}

	/**
	 * Gets test order details
	 * 
	 * @param TestOrderId Represents test order id
	 * 
	 */
	private static MailMergeDataTable getTestOrderDetails(int TestOrderId) throws Exception {
		// Loads the XML file.
		File file = new File(getDataDir("TestOrderDetails.xml"));
		// Create a new instance for the JAXBContext.
		JAXBContext jaxbContext = JAXBContext.newInstance(TestOrderDetails.class);
		// Reads the XML file.
		Unmarshaller jaxbUnmarshaller = jaxbContext.createUnmarshaller();
		TestOrderDetails testOrders = (TestOrderDetails) jaxbUnmarshaller.unmarshal(file);
		// Gets the list of test order details.
		List<TestOrderDetail> list = testOrders.getTestOrderDetails();
		ListSupport<TestOrderDetail> testOrderList = new ListSupport<TestOrderDetail>(TestOrderDetail.class);
		for (TestOrderDetail testOrder : list) {
			if (testOrder.getOrderID().equals(Int32Support.toString(TestOrderId))) {
				testOrderList.add(testOrder);
			}
		}
		// Creates an instance of MailMergeDataTable by specifying mail merge group name and IEnumerable collection.
		MailMergeDataTable dataTable = new MailMergeDataTable("Order", testOrderList);
		return dataTable;
	}

	/**
	 * Gets test order totals
	 * 
	 * @param TestOrderId Represents test order id
	 * 
	 */
	private static MailMergeDataTable getTestOrderTotals(int TestOrderId) throws Exception {
		// Loads the XML file.
		File file = new File(getDataDir("OrderTotals.xml"));
		// Create a new instance for the JAXBContext.
		JAXBContext jaxbContext = JAXBContext.newInstance(TestOrderTotals.class);
		// Reads the XML file.
		Unmarshaller jaxbUnmarshaller = jaxbContext.createUnmarshaller();
		TestOrderTotals testOrders = (TestOrderTotals) jaxbUnmarshaller.unmarshal(file);
		// Gets the list of test order details.
		List<TestOrderTotal> list = testOrders.getTestOrderTotals();
		ListSupport<TestOrderTotal> testOrderList = new ListSupport<TestOrderTotal>(TestOrderTotal.class);
		for (TestOrderTotal testOrder : list) {
			if (testOrder.getOrderID().equals(Int32Support.toString(TestOrderId))) {
				testOrderList.add(testOrder);
			}
		}
		// Creates an instance of MailMergeDataTable by specifying mail merge group name and IEnumerable collection.
		MailMergeDataTable dataTable = new MailMergeDataTable("OrderTotals", testOrderList);
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
