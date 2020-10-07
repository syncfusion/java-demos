package employeereport;

public class Employees {
	private String m_employeeID;
	private String m_lastName;
	private String m_firstName;
	private String m_title;
	private String m_titleOfCourtesy;
	private String m_birthDate;
	private String m_hireDate;
	private String m_address;
	private String m_city;
	private String m_region;
	private String m_postalCode;
	private String m_country;
	private String m_homePhone;
	private String m_extension;
	private String m_photo;
	private String m_notes;
	private String m_reportsTo;

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

	public String getTitle() throws Exception {
		return m_title;
	}

	public String setTitle(String value) throws Exception {
		m_title = value;
		return value;
	}

	public String getTitleOfCourtesy() throws Exception {
		return m_titleOfCourtesy;
	}

	public String setTitleOfCourtesy(String value) throws Exception {
		m_titleOfCourtesy = value;
		return value;
	}

	public String getBirthDate() throws Exception {
		return m_birthDate;
	}

	public String setBirthDate(String value) throws Exception {
		m_birthDate = value;
		return value;
	}

	public String getHireDate() throws Exception {
		return m_hireDate;
	}

	public String setHireDate(String value) throws Exception {
		m_hireDate = value;
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

	public String getCountry() throws Exception {
		return m_country;
	}

	public String setCountry(String value) throws Exception {
		m_country = value;
		return value;
	}

	public String getHomePhone() throws Exception {
		return m_homePhone;
	}

	public String setHomePhone(String value) throws Exception {
		m_homePhone = value;
		return value;
	}

	public String getExtension() throws Exception {
		return m_extension;
	}

	public String setExtension(String value) throws Exception {
		m_extension = value;
		return value;
	}

	public String getPhoto() throws Exception {
		return m_photo;
	}

	public String setPhoto(String value) throws Exception {
		m_photo = value;
		return value;
	}

	public String getNotes() throws Exception {
		return m_notes;
	}

	public String setNotes(String value) throws Exception {
		m_notes = value;
		return value;
	}

	public String getReportsTo() throws Exception {
		return m_reportsTo;
	}

	public String setReportsTo(String value) throws Exception {
		m_reportsTo = value;
		return value;
	}

	public Employees(String employeeID, String lastName, String firstName, String title, String titleOfCourtesy,
			String birthDate, String hireDate, String address, String city, String region, String postalCode,
			String country, String homePhone, String extension, String photo, String notes, String reportsTo)
			throws Exception {
		m_employeeID = employeeID;
		m_lastName = lastName;
		m_firstName = firstName;
		m_title = title;
		m_titleOfCourtesy = titleOfCourtesy;
		m_birthDate = birthDate;
		m_hireDate = hireDate;
		m_address = address;
		m_city = city;
		m_region = region;
		m_postalCode = postalCode;
		m_country = country;
		m_homePhone = homePhone;
		m_extension = extension;
		m_photo = photo;
		m_notes = notes;
		m_reportsTo = reportsTo;
	}

	public Employees() throws Exception {
	}
}
