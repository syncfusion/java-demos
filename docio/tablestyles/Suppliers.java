package tablestyles;

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

	public String getSupplierID() throws Exception {
		return m_id;
	}

	public String setSupplierID(String value) throws Exception {
		m_id = value;
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
		return m_contactName;
	}

	public String setContactName(String value) throws Exception {
		m_contactName = value;
		return value;
	}

	public String getContactTitle() throws Exception {
		return m_contactTitle;
	}

	public String setContactTitle(String value) throws Exception {
		m_contactTitle = value;
		return value;
	}

	public String getRegion() throws Exception {
		return m_region;
	}

	public String setRegion(String value) throws Exception {
		m_region = value;
		return value;
	}

	public String getPostalCode() throws Exception {
		return m_postalCode;
	}

	public String setPostalCode(String value) throws Exception {
		m_postalCode = value;
		return value;
	}

	public String getPhone() throws Exception {
		return m_phone;
	}

	public String setPhone(String value) throws Exception {
		m_phone = value;
		return value;
	}

	public String getFax() throws Exception {
		return m_fax;
	}

	public String setFax(String value) throws Exception {
		m_fax = value;
		return value;
	}

	public String getHomePage() throws Exception {
		return m_homePage;
	}

	public String setHomePage(String value) throws Exception {
		m_homePage = value;
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

	public Suppliers(String id, String companyName, String contantName, String address, String city, String country)
			throws Exception {
		m_id = id;
		m_companyName = companyName;
		m_contactName = contantName;
		m_address = address;
		m_city = city;
		m_country = country;
	}

	public Suppliers() throws Exception {
	}
}
