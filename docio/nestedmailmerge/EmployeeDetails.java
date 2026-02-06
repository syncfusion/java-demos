package nestedmailmerge;

import java.util.List;

import javax.xml.bind.annotation.XmlElement;

public class EmployeeDetails {
	private String m_employeeID;
	private String m_lastName;
	private String m_firstName;
	private String m_address;
	private String m_city;
	private String m_country;
	private String m_extension;
	private List<CustomerDetails> m_customerDetails;

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
	 * Gets the last name of the employee.
	 */
	@XmlElement(name = "LastName")
	public String getLastName() throws Exception {
		return m_lastName;
	}

	/**
	 * Sets the last name of the employee.
	 */
	public String setLastName(String value) throws Exception {
		m_lastName = value;
		return value;
	}

	/**
	 * Gets the first name of the employee.
	 */
	@XmlElement(name = "FirstName")
	public String getFirstName() throws Exception {
		return m_firstName;
	}

	/**
	 * Sets the first name of the employee.
	 */
	public String setFirstName(String value) throws Exception {
		m_firstName = value;
		return value;
	}

	/**
	 * Gets the address of the employee.
	 */
	@XmlElement(name = "Address")
	public String getAddress() throws Exception {
		return m_address;
	}

	/**
	 * Sets the address of the employee.
	 */
	public String setAddress(String value) throws Exception {
		m_address = value;
		return value;
	}

	/**
	 * Gets the city of the employee.
	 */
	@XmlElement(name = "City")
	public String getCity() throws Exception {
		return m_city;
	}

	/**
	 * Sets the city of the employee.
	 */
	public String setCity(String value) throws Exception {
		m_city = value;
		return value;
	}

	/**
	 * Gets the country of the employee.
	 */
	@XmlElement(name = "Country")
	public String getCountry() throws Exception {
		return m_country;
	}

	/**
	 * Sets the country of the employee.
	 */
	public String setCountry(String value) throws Exception {
		m_country = value;
		return value;
	}

	/**
	 * Gets the extension of the employee.
	 */
	@XmlElement(name = "Extension")
	public String getExtension() throws Exception {
		return m_extension;
	}

	/**
	 * Sets the extension of the employee.
	 */
	public String setExtension(String value) throws Exception {
		m_extension = value;
		return value;

	}

	/**
	 * Gets the customer details.
	 */
	@XmlElement(name = "Customers")
	public List<CustomerDetails> getCustomerDetails() {
		return m_customerDetails;
	}

	/**
	 * Sets the customer details.
	 */
	public void setCustomerDetails(List<CustomerDetails> customerDetails) {
		m_customerDetails = customerDetails;
	}
}
