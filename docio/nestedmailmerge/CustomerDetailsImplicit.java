package nestedmailmerge;

import com.syncfusion.javahelper.system.collections.generic.ListSupport;

public class CustomerDetailsImplicit {
	private String m_employeeID;
	private String m_customerID;
	private String m_companyName;
	private String m_city;
	private String m_country;
	private ListSupport<OrderDetails> m_orders;

	public ListSupport<OrderDetails> getOrders() throws Exception {
		if (m_orders == null)
			m_orders = new ListSupport<OrderDetails>(OrderDetails.class);
		return m_orders;
	}

	public ListSupport<OrderDetails> setOrders(ListSupport<OrderDetails> value) throws Exception {
		m_orders = value;
		return value;
	}

	public String getEmployeeID() throws Exception {
		return m_employeeID;
	}

	public String setEmployeeID(String value) throws Exception {
		m_employeeID = value;
		return value;
	}

	public String getCustomerID() throws Exception {
		return m_customerID;
	}

	public String setCustomerID(String value) throws Exception {
		m_customerID = value;
		return value;
	}

	public String getCompanyName() throws Exception {
		return m_companyName;
	}

	public String setCompanyName(String value) throws Exception {
		m_companyName = value;
		return value;
	}

	public String getContactName() throws Exception {
		return m_companyName;
	}

	public String setContactName(String value) throws Exception {
		m_companyName = value;
		return value;
	}

	public String getCity() throws Exception {
		return m_city;
	}

	public String setCity(String value) throws Exception {
		m_city = value;
		return value;
	}

	public String getCountry() throws Exception {
		return m_country;
	}

	public String setCountry(String value) throws Exception {
		m_country = value;
		return value;
	}

	public CustomerDetailsImplicit() throws Exception {
		m_orders = new ListSupport<OrderDetails>(OrderDetails.class);
	}
}
