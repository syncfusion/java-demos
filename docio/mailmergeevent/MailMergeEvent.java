package mailmergeevent;

import java.io.*;
import mailmergeevent.ProductDetail;
import mailmergeevent.Product_PriceList;
import com.syncfusion.docio.*;
import com.syncfusion.javahelper.system.*;
import com.syncfusion.javahelper.system.collections.generic.ListSupport;
import com.syncfusion.javahelper.system.drawing.ColorSupport;
import com.syncfusion.javahelper.system.io.*;
import com.syncfusion.javahelper.system.xml.*;

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
		FormatType type = FormatType.Docx;
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
		// Gets the product_pricelist as “IEnumerable” collection.
		ListSupport<Product_PriceList> product_PriceList = new ListSupport<Product_PriceList>(Product_PriceList.class);
		FileStreamSupport fs = new FileStreamSupport(getDataDir("ProductPriceList.xml"), FileMode.Open,
				FileAccess.Read);
		XmlReaderSupport reader = XmlReaderSupport.create(fs);
		if (reader == null)
			throw new Exception("reader");
		while (reader.getNodeType().getEnumValue() != XmlNodeType.Element.getEnumValue())
			reader.read();
		if (!(reader.getLocalName() == "Products"))
			throw new Exception("Unexpected xml tag " + reader.getLocalName());
		reader.read();
		while (reader.getNodeType().getEnumValue() == XmlNodeType.Whitespace.getEnumValue())
			reader.read();
		while (!(reader.getLocalName() == "Products")) {
			if (reader.getNodeType().getEnumValue() == XmlNodeType.Element.getEnumValue()) {
				switch ((reader.getLocalName()) == null ? "string_null_value" : (reader.getLocalName())) {
				case "Product_PriceList":

					product_PriceList.add(getProduct_PriceList(reader));
					break;
				}
			} else

			{
				reader.read();
				if (((reader.getLocalName() == "Products"))
						&& reader.getNodeType().getEnumValue() == XmlNodeType.EndElement.getEnumValue())
					break;
			}
		}
		// Creates an instance of “MailMergeDataTable” by specifying mail merge group name and “IEnumerable” collection.
		MailMergeDataTable dataTable1 = new MailMergeDataTable("Product_PriceList", product_PriceList);
		reader.close();
		fs.close();
		return dataTable1;
	}

	/**
	 * 
	 * Gets the product price list.
	 * 
	 * @param reader The reader.
	 */
	private static Product_PriceList getProduct_PriceList(XmlReaderSupport reader) throws Exception {
		if (reader == null)
			throw new Exception("reader");
		while (reader.getNodeType().getEnumValue() != XmlNodeType.Element.getEnumValue())
			reader.read();
		if (!(reader.getLocalName() == "Product_PriceList"))
			throw new Exception("Unexpected xml tag " + reader.getLocalName());
		reader.read();
		while (reader.getNodeType().getEnumValue() == XmlNodeType.Whitespace.getEnumValue())
			reader.read();
		Product_PriceList product_PriceList = new Product_PriceList();
		while (!(reader.getLocalName() == "Product_PriceList")) {
			if (reader.getNodeType().getEnumValue() == XmlNodeType.Element.getEnumValue()) {
				switch ((reader.getLocalName()) == null ? "string_null_value" : (reader.getLocalName())) {
				case "ProductName":

					product_PriceList.setProductName(reader.readContentAsString());
					break;
				case "Price":

					product_PriceList.setPrice(reader.readContentAsString());
					break;
				default:

					reader.skip();
					break;
				}
			} else

			{
				reader.read();
				if (((reader.getLocalName() == "Product_PriceList"))
						&& reader.getNodeType().getEnumValue() == XmlNodeType.EndElement.getEnumValue())
					break;
			}
		}
		return product_PriceList;
	}

	/**
	 * 
	 * Gets the mail merge data table.
	 * 
	 */
	private static MailMergeDataTable getMailMergeDataTableProductData() throws Exception {
		// Gets the product detail as “IEnumerable” collection.
		ListSupport<ProductDetail> productDetail = new ListSupport<ProductDetail>(ProductDetail.class);
		FileStreamSupport fs = new FileStreamSupport(getDataDir("Product.xml"), FileMode.Open, FileAccess.Read);
		XmlReaderSupport reader = XmlReaderSupport.create(fs);
		if (reader == null)
			throw new Exception("reader");
		while (reader.getNodeType().getEnumValue() != XmlNodeType.Element.getEnumValue())
			reader.read();
		if (!(reader.getLocalName() == "ProductList"))
			throw new Exception("Unexpected xml tag " + reader.getLocalName());
		reader.read();
		while (reader.getNodeType().getEnumValue() == XmlNodeType.Whitespace.getEnumValue())
			reader.read();
		while (!(reader.getLocalName() == "ProductList")) {
			if (reader.getNodeType().getEnumValue() == XmlNodeType.Element.getEnumValue()) {
				switch ((reader.getLocalName()) == null ? "string_null_value" : (reader.getLocalName())) {
				case "Products":

					productDetail.add(getProductDetail(reader));
					break;
				}
			} else

			{
				reader.read();
				if (((reader.getLocalName() == "ProductList"))
						&& reader.getNodeType().getEnumValue() == XmlNodeType.EndElement.getEnumValue())
					break;
			}
		}
		// Creates an instance of “MailMergeDataTable” by specifying mail merge group name and “IEnumerable” collection.
		MailMergeDataTable dataTable2 = new MailMergeDataTable("ProductDetail", productDetail);
		reader.close();
		fs.close();
		return dataTable2;
	}

	/**
	 * 
	 * Gets the product details.
	 * 
	 * @param reader The reader.
	 */
	private static ProductDetail getProductDetail(XmlReaderSupport reader) throws Exception {
		if (reader == null)
			throw new Exception("reader");
		while (reader.getNodeType().getEnumValue() != XmlNodeType.Element.getEnumValue())
			reader.read();
		if (!(reader.getLocalName() == "Products"))
			throw new Exception("Unexpected xml tag " + reader.getLocalName());
		reader.read();
		while (reader.getNodeType().getEnumValue() == XmlNodeType.Whitespace.getEnumValue())
			reader.read();
		ProductDetail productDetail = new ProductDetail();
		while (!(reader.getLocalName() == "Products")) {
			if (reader.getNodeType().getEnumValue() == XmlNodeType.Element.getEnumValue()) {
				switch ((reader.getLocalName()) == null ? "string_null_value" : (reader.getLocalName())) {
				case "SNO":

					productDetail.setSNO(reader.readContentAsString());
					break;
				case "ProductName":

					productDetail.setProductName(reader.readContentAsString());
					break;
				case "ProductImage":

					productDetail.setProductImage(reader.readContentAsString());
					break;
				default:

					reader.skip();
					break;
				}
			} else

			{
				reader.read();
				if ((reader.getLocalName() == "Products")
						&& reader.getNodeType().getEnumValue() == XmlNodeType.EndElement.getEnumValue())
					break;
			}
		}
		return productDetail;
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
