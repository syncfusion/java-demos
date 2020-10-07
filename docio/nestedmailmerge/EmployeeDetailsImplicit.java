package nestedmailmerge;

import com.syncfusion.javahelper.system.collections.generic.ListSupport;

public class EmployeeDetailsImplicit {
	private String m_employeeID;
	private String m_lastName;
	private String m_firstName;
	private String m_address;
	private String m_city;
	private String m_country;
	private String m_extension;
	private ListSupport<CustomerDetailsImplicit> m_customers;

	public String getEmployeeID() throws Exception {
		return m_employeeID;
	}

	public String setEmployeeID(String value) throws Exception {
		m_employeeID = value;
		return value;
	}

	public String getLastName() throws Exception {
		return m_lastName;
	}

	public String setLastName(String value) throws Exception {
		m_lastName = value;
		return value;
	}

	public String getFirstName() throws Exception {
		return m_firstName;
	}

	public String setFirstName(String value) throws Exception {
		m_firstName = value;
		return value;
	}

	public String getAddress() throws Exception {
		return m_address;
	}

	public String setAddress(String value) throws Exception {
		m_address = value;
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

	public String getExtension() throws Exception {
		return m_extension;
	}

	public String setExtension(String value) throws Exception {
		m_extension = value;
		return value;
	}

	public ListSupport<CustomerDetailsImplicit> getCustomers() throws Exception {
		if (m_customers == null)
			m_customers = new ListSupport<CustomerDetailsImplicit>(CustomerDetailsImplicit.class);
		return m_customers;
	}

	public ListSupport<CustomerDetailsImplicit> setCustomers(ListSupport<CustomerDetailsImplicit> value)
			throws Exception {
		m_customers = value;
		return value;
	}

	public EmployeeDetailsImplicit() throws Exception {
		m_customers = new ListSupport<CustomerDetailsImplicit>(CustomerDetailsImplicit.class);
	}
}
