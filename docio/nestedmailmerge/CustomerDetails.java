package nestedmailmerge;

import java.util.List;

import javax.xml.bind.annotation.*;

@XmlRootElement(name = "Customers")
public class CustomerDetails {
	private String m_employeeID;
	private String m_customerID;
	private String m_companyName;
	private String m_contactName;
	private String m_city;
	private String m_country;
	private List<OrderDetails> m_orders;

	/**
	 * Gets the Id of the employee.
	 */
	@XmlElement(name = "EmployeeID")
	public String getEmployeeID() throws Exception {
		return m_employeeID;
	}

	/**
	 * Sets the Id of the employee.
	 */
	public String setEmployeeID(String value) throws Exception {
		m_employeeID = value;
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
	 * Gets the name of the company.
	 */
	@XmlElement(name = "CompanyName")
	public String getCompanyName() throws Exception {
		return m_companyName;
	}

	/**
	 * Sets the name of the company.
	 */
	public String setCompanyName(String value) throws Exception {
		m_companyName = value;
		return value;
	}

	/**
	 * Gets the contact name of the customer.
	 */
	@XmlElement(name = "ContactName")
	public String getContactName() throws Exception {
		return m_contactName;
	}

	/**
	 * Sets the contact name of the customer.
	 */
	public String setContactName(String value) throws Exception {
		m_contactName = value;
		return value;
	}

	/**
	 * Gets the city of the customer.
	 */
	@XmlElement(name = "City")
	public String getCity() throws Exception {
		return m_city;
	}

	/**
	 * Sets the city of the customer.
	 */
	public String setCity(String value) throws Exception {
		m_city = value;
		return value;
	}

	/**
	 * Gets the country of the customer.
	 */
	@XmlElement(name = "Country")
	public String getCountry() throws Exception {
		return m_country;
	}

	/**
	 * Sets the country of the customer.
	 */
	public String setCountry(String value) throws Exception {
		m_country = value;
		return value;
	}

	/**
	 * Gets the order details of the customer.
	 */
	@XmlElement(name = "Orders")
	public List<OrderDetails> getOrders() throws Exception {
		return m_orders;
	}

	/**
	 * Sets the order details of the customer.
	 */
	public void setOrders(List<OrderDetails> orders) throws Exception {
		m_orders = orders;
		this.m_orders = orders;
	}

}
