package mailmergeevent;

import javax.xml.bind.annotation.XmlElement;

public class Product_PriceList {
	private String m_productName;
	private String m_price;

	/**
	 * Gets the name of the product.
	 */
	@XmlElement(name = "ProductName")
	public String getProductName() throws Exception {
		return m_productName;
	}

	/**
	 * Sets the name of the product.
	 */
	public String setProductName(String value) throws Exception {
		m_productName = value;
		return value;
	}

	/**
	 * Gets the price of the product.
	 */
	@XmlElement(name = "Price")
	public String getPrice() throws Exception {
		return m_price;
	}

	/**
	 * Sets the price of the product.
	 */
	public String setPrice(String value) throws Exception {
		m_price = value;
		return value;
	}
}
