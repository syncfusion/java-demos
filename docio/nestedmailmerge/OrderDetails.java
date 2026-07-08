package nestedmailmerge;

import javax.xml.bind.annotation.XmlElement;

public class OrderDetails {
	private String m_orderID;
	private String m_customerID;
	private String m_orderDate;
	private String m_requiredDate;
	private String m_shippedDate;

	/**
	 * Gets the Id of the order.
	 */
	@XmlElement(name = "OrderID")
	public String getOrderID() throws Exception {
		return m_orderID;
	}

	/**
	 * Sets the Id of the order.
	 */
	public String setOrderID(String value) throws Exception {
		m_orderID = value;
		return value;
	}

	/**
	 * Gets the Id of the customer.
	 */
	@XmlElement(name = "CustomerID")
	public String getCustomerID() throws Exception {
		return m_customerID;
	}

	/**
	 * Sets the Id of the customer.
	 */
	public String setCustomerID(String value) throws Exception {
		m_customerID = value;
		return value;
	}

	/**
	 * Gets the date of order.
	 */
	@XmlElement(name = "OrderDate")
	public String getOrderDate() throws Exception {
		return m_orderDate;
	}

	/**
	 * Sets the date of order.
	 */
	public String setOrderDate(String value) throws Exception {
		m_orderDate = value;
		return value;
	}

	/**
	 * Gets the required date of order.
	 */
	@XmlElement(name = "RequiredDate")
	public String getRequiredDate() throws Exception {
		return m_requiredDate;
	}

	/**
	 * Sets the required date of order.
	 */
	public String setRequiredDate(String value) throws Exception {
		m_requiredDate = value;
		return value;
	}

	/**
	 * Gets the shipped date of order.
	 */
	@XmlElement(name = "ShippedDate")
	public String getShippedDate() throws Exception {
		return m_shippedDate;
	}

	/**
	 * Sets the shipped date of order.
	 */
	public String setShippedDate(String value) throws Exception {
		m_shippedDate = value;
		return value;
	}
}
