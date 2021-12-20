package mailmergeevent;

import java.io.*;
import java.util.List;
import javax.xml.bind.JAXBContext;
import javax.xml.bind.Unmarshaller;
import com.syncfusion.docio.*;
import com.syncfusion.javahelper.system.collections.generic.ListSupport;
import com.syncfusion.javahelper.system.drawing.ColorSupport;
import com.syncfusion.javahelper.system.io.*;

public class MailMergeEvent {
	public static void main(String[] args) throws Exception {
		// Loads the template.
		WordDocument document = new WordDocument(getDataDir("MailMergeEventTemplate.docx"), FormatType.Docx);
		// Using Merge events to do conditional formatting during runtime.
		document.getMailMerge().MergeField.add("alternateRows_MergeField", new MergeFieldEventHandler() {
			ListSupport<MergeFieldEventHandler> delegateList = new ListSupport<MergeFieldEventHandler>(
					MergeFieldEventHandler.class);
			// Represents event handling for MergeFieldEventHandlerCollection.
			public void invoke(Object sender, MergeFieldEventArgs args) throws Exception {
				alternateRows_MergeField(sender, args);
			}
			// Represents the method that handles MergeField event.
			public void dynamicInvoke(Object... args) throws Exception {
				alternateRows_MergeField((Object) args[0], (MergeFieldEventArgs) args[1]);
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
		// Using Merge events to do conditional formatting during runtime.
		document.getMailMerge().MergeImageField.add("mergeField_ProductImage", new MergeImageFieldEventHandler() {
			ListSupport<MergeImageFieldEventHandler> delegateList = new ListSupport<MergeImageFieldEventHandler>(
					MergeImageFieldEventHandler.class);
			// Represents event handling for MergeFieldEventHandlerCollection.
			public void invoke(Object sender, MergeImageFieldEventArgs args) throws Exception {
				mergeField_ProductImage(sender, args);
			}
			// Represents the method that handles MergeField event.
			public void dynamicInvoke(Object... args) throws Exception {
				mergeField_ProductImage((Object) args[0], (MergeImageFieldEventArgs) args[1]);
			}
			// Represents the method that handles MergeField event to add collection item.
			public void add(MergeImageFieldEventHandler delegate) throws Exception {
				if (delegate != null)
					delegateList.add(delegate);
			}
			// Represents the method that handles MergeField event to remove collection item.
			public void remove(MergeImageFieldEventHandler delegate) throws Exception {
				if (delegate != null)
					delegateList.remove(delegate);
			}
		});
		// Gets the mail merge data table.
		MailMergeDataTable mailMergeDataTablePriceList = getMailMergeDataTablePriceList();
		MailMergeDataTable mailMergeDataTableProductData = getMailMergeDataTableProductData();
		// Gets mail merge from word document.
		MailMerge mailMerge = document.getMailMerge();
		// Execute Mail Merge with groups.
		mailMerge.executeGroup(mailMergeDataTablePriceList);
		mailMerge.executeGroup(mailMergeDataTableProductData);
		// Save and close the document.
		document.save("Mail Merge Event.docx", FormatType.Docx);
		document.close();
		System.out.println("Word document generated successfully.");
	}

	/**
	 * 
	 * Get alternateRows for MergeField
	 * 
	 * @param sender Represents the object.
	 * @param args   Represents MergeField EventArguments
	 */
	private static void alternateRows_MergeField(Object sender, MergeFieldEventArgs args) throws Exception {
		if (args.getRowIndex() % 2 == 0) {
			// Sets text color to the merge field event.
			args.getCharacterFormat().setTextColor(ColorSupport.fromArgb(255, 102, 0));
		}
	}

	/**
	 * 
	 * Get mergeField for ProductImage
	 * 
	 * @param sender Represents the object.
	 * @param args   Represents MergeImageField EventArguments
	 */
	private static void mergeField_ProductImage(Object sender, MergeImageFieldEventArgs args) throws Exception {
		 // Gets the image from disk during Merge.
		if (args.getFieldName().equals("ProductImage")) {
			String ProductFileName = args.getFieldValue().toString();
			FileStreamSupport fs = new FileStreamSupport(getDataDir(ProductFileName), FileMode.Open,
					FileAccess.Read);
			ByteArrayInputStream stream = new ByteArrayInputStream(fs.toArray());
			args.setImageStream(stream);
		}
	}

	/**
	 * 
	 * Gets the mail merge data table.
	 * 
	 */
	private static MailMergeDataTable getMailMergeDataTablePriceList() throws Exception {
		// Loads the XML file.
		File file = new File(getDataDir("ProductPriceList.xml"));
		// Create a new instance for the JAXBContext.
		JAXBContext jaxbContext = JAXBContext.newInstance(ProductDetail.class);
		// Reads the XML file.
		Unmarshaller jaxbUnmarshaller = jaxbContext.createUnmarshaller();
		ProductDetail products = (ProductDetail) jaxbUnmarshaller.unmarshal(file);
		// Gets the list of product price details.
		List<Product_PriceList> list = products.getPriceList();
		ListSupport<Product_PriceList> priceList = new ListSupport<Product_PriceList>(Product_PriceList.class);
		priceList.addRange(list);
		// Creates an instance of MailMergeDataTable by specifying mail merge group name and IEnumerable collection.
		MailMergeDataTable dataTable = new MailMergeDataTable("Product_PriceList", priceList);
		return dataTable;
	}

	/**
	 * 
	 * Gets the mail merge data table.
	 * 
	 */
	private static MailMergeDataTable getMailMergeDataTableProductData() throws Exception {
		// Loads the XML file.
		File file = new File(getDataDir("Product.xml"));
		// Create a new instance for the JAXBContext.
		JAXBContext jaxbContext = JAXBContext.newInstance(ProductList.class);
		// Reads the XML file.
		Unmarshaller jaxbUnmarshaller = jaxbContext.createUnmarshaller();
		ProductList products = (ProductList) jaxbUnmarshaller.unmarshal(file);
		// Gets the list of product details.
		List<Products> list = products.getProducts();
		ListSupport<Products> productList = new ListSupport<Products>(Products.class);
		productList.addRange(list);
		// Creates an instance of MailMergeDataTable by specifying mail merge group name and IEnumerable collection.
		MailMergeDataTable dataTable = new MailMergeDataTable("ProductDetail", productList);
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
