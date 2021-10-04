package salesinvoice;

import javax.xml.bind.annotation.XmlElement;

public class TestOrderDetail {
	private String m_orderID;
	private String m_productID;
	private String m_productName;
	private String m_unitPrice;
	private String m_quantity;
	private String m_discount;
	private String m_extendedPrice;

	/**
	 * Gets the ID of the order.
	 */
	@XmlElement(name = "OrderID")
	public String getOrderID() throws Exception {
		return m_orderID;
	}

	/**
	 * Sets the ID of the order.
	 */
	public String setOrderID(String value) throws Exception {
		m_orderID = value;
		return value;
	}

	/**
	 * Gets the ID of the product.
	 */
	@XmlElement(name = "ProductID")
	public String getProductID() throws Exception {
		return m_productID;
	}

	/**
	 * Sets the ID of the product.
	 */
	public String setProductID(String value) throws Exception {
		m_productID = value;
		return value;
	}

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
	 * Gets the unit price of the product.
	 */
	@XmlElement(name = "UnitPrice")
	public String getUnitPrice() throws Exception {
		return m_unitPrice;
	}

	/**
	 * Sets the unit price of the product.
	 */
	public String setUnitPrice(String value) throws Exception {
		m_unitPrice = value;
		return value;
	}

	/**
	 * Gets the quantity of the product.
	 */
	@XmlElement(name = "Quantity")
	public String getQuantity() throws Exception {
		return m_quantity;
	}

	/**
	 * Sets the quantity of the product.
	 */
	public String setQuantity(String value) throws Exception {
		m_quantity = value;
		return value;
	}

	/**
	 * Gets the discount for the product.
	 */
	@XmlElement(name = "Discount")
	public String getDiscount() throws Exception {
		return m_discount;
	}

	/**
	 * Sets the discount for the product.
	 */
	public String setDiscount(String value) throws Exception {
		m_discount = value;
		return value;
	}

	/**
	 * Gets the extended price of the product.
	 */
	@XmlElement(name = "ExtendedPrice")
	public String getExtendedPrice() throws Exception {
		return m_extendedPrice;
	}

	/**
	 * Sets the extended price of the product.
	 */
	public String setExtendedPrice(String value) throws Exception {
		m_extendedPrice = value;
		return value;
	}

}
