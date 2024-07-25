package tablestyles;

import javax.xml.bind.annotation.XmlElement;

public class Suppliers {
	private String m_id;
	private String m_companyName;
	private String m_contactName;
	private String m_address;
	private String m_city;
	private String m_country;
	private String m_contactTitle;
	private String m_region;
	private String m_postalCode;
	private String m_phone;
	private String m_fax;
	private String m_homePage;

	/**
	 * Gets the Id of the supplier.
	 */
	@XmlElement(name = "SupplierID")
	public String getSupplierID() throws Exception {
		return m_id;
	}

	/**
	 * Sets the Id of the supplier.
	 */
	public String setSupplierID(String value) throws Exception {
		m_id = value;
		return value;
	}

	/**
	 * Gets the company name of the supplier.
	 */
	@XmlElement(name = "CompanyName")
	public String getCompanyName() throws Exception {
		return m_companyName;
	}

	/**
	 * Sets the company name of the supplier.
	 */
	public String setCompanyName(String value) throws Exception {
		m_companyName = value;
		return value;
	}

	/**
	 * Gets the contact name of the supplier.
	 */
	@XmlElement(name = "ContactName")
	public String getContactName() throws Exception {
		return m_contactName;
	}

	/**
	 * Sets the contact name of the supplier.
	 */
	public String setContactName(String value) throws Exception {
		m_contactName = value;
		return value;
	}

	/**
	 * Gets the contact title of the supplier.
	 */
	@XmlElement(name = "ContactTitle")
	public String getContactTitle() throws Exception {
		return m_contactTitle;
	}

	/**
	 * Sets the contact title of the supplier.
	 */
	public String setContactTitle(String value) throws Exception {
		m_contactTitle = value;
		return value;
	}

	/**
	 * Gets the region.
	 */
	@XmlElement(name = "Region")
	public String getRegion() throws Exception {
		return m_region;
	}

	/**
	 * Sets the region.
	 */
	public String setRegion(String value) throws Exception {
		m_region = value;
		return value;
	}

	/**
	 * Gets the postal code.
	 */
	@XmlElement(name = "PostalCode")
	public String getPostalCode() throws Exception {
		return m_postalCode;
	}

	/**
	 * Sets the postal code.
	 */
	public String setPostalCode(String value) throws Exception {
		m_postalCode = value;
		return value;
	}

	/**
	 * Gets the contact number.
	 */
	@XmlElement(name = "Phone")
	public String getPhone() throws Exception {
		return m_phone;
	}

	/**
	 * Sets the contact number.
	 */
	public String setPhone(String value) throws Exception {
		m_phone = value;
		return value;
	}

	/**
	 * Gets the fax details.
	 */
	@XmlElement(name = "Fax")
	public String getFax() throws Exception {
		return m_fax;
	}

	/**
	 * Sets the fax details.
	 */
	public String setFax(String value) throws Exception {
		m_fax = value;
		return value;
	}

	/**
	 * Gets the home page.
	 */
	@XmlElement(name = "HomePage")
	public String getHomePage() throws Exception {
		return m_homePage;
	}

	/**
	 * Sets the home page.
	 */
	public String setHomePage(String value) throws Exception {
		m_homePage = value;
		return value;
	}

	/**
	 * Gets the address.
	 */
	@XmlElement(name = "Address")
	public String getAddress() throws Exception {
		return m_address;
	}

	/**
	 * Sets the address.
	 */
	public String setAddress(String value) throws Exception {
		m_address = value;
		return value;
	}

	/**
	 * Gets the city.
	 */
	@XmlElement(name = "City")
	public String getCity() throws Exception {
		return m_city;
	}

	/**
	 * Sets the city.
	 */
	public String setCity(String value) throws Exception {
		m_city = value;
		return value;
	}

	/**
	 * Gets the country.
	 */
	@XmlElement(name = "Country")
	public String getCountry() throws Exception {
		return m_country;
	}

	/**
	 * Sets the country.
	 */
	public String setCountry(String value) throws Exception {
		m_country = value;
		return value;
	}
}
