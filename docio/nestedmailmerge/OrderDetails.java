package nestedmailmerge;

public class OrderDetails {
	private String m_orderID;
	private String m_customerID;
	private String m_orderDate;
	private String m_requiredDate;
	private String m_shippedDate;

	public String getOrderID() throws Exception {
		return m_orderID;
	}

	public String setOrderID(String value) throws Exception {
		m_orderID = value;
		return value;
	}

	public String getCustomerID() throws Exception {
		return m_customerID;
	}

	public String setCustomerID(String value) throws Exception {
		m_customerID = value;
		return value;
	}

	public String getOrderDate() throws Exception {
		return m_orderDate;
	}

	public String setOrderDate(String value) throws Exception {
		m_orderDate = value;
		return value;
	}

	public String getRequiredDate() throws Exception {
		return m_requiredDate;
	}

	public String setRequiredDate(String value) throws Exception {
		m_requiredDate = value;
		return value;
	}

	public String getShippedDate() throws Exception {
		return m_shippedDate;
	}

	public String setShippedDate(String value) throws Exception {
		m_shippedDate = value;
		return value;
	}
}
