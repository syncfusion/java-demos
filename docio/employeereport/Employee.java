package employeereport;

import javax.xml.bind.annotation.XmlElement;

public class Employee {
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

	/**
	 * Gets the employee id.
	 */
	@XmlElement(name = "EmployeeID")
	public String getEmployeeID() throws Exception {
		return m_employeeID;
	}

	/**
	 * Sets the employee id.
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
	 * Gets the title of the employee.
	 */
	@XmlElement(name = "Title")
	public String getTitle() throws Exception {
		return m_title;
	}

	/**
	 * Sets the title of the employee.
	 */
	public String setTitle(String value) throws Exception {
		m_title = value;
		return value;
	}

	/**
	 * Gets the title of courtesy of the employee.
	 */
	@XmlElement(name = "TitleOfCourtesy")
	public String getTitleOfCourtesy() throws Exception {
		return m_titleOfCourtesy;
	}

	/**
	 * Sets the title of courtesy of the employee.
	 */
	public String setTitleOfCourtesy(String value) throws Exception {
		m_titleOfCourtesy = value;
		return value;
	}

	/**
	 * Gets the birth date of the employee.
	 */
	@XmlElement(name = "BirthDate")
	public String getBirthDate() throws Exception {
		return m_birthDate;
	}

	/**
	 * Sets the birth date of the employee.
	 */
	public String setBirthDate(String value) throws Exception {
		m_birthDate = value;
		return value;
	}

	/**
	 * Gets the hire date of the employee.
	 */
	@XmlElement(name = "HireDate")
	public String getHireDate() throws Exception {
		return m_hireDate;
	}

	/**
	 * Sets the hire date of the employee.
	 */
	public String setHireDate(String value) throws Exception {
		m_hireDate = value;
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
	 * Gets the region of the employee.
	 */
	@XmlElement(name = "Region")
	public String getRegion() throws Exception {
		return m_region;
	}

	/**
	 * Sets the hire date of the employee.
	 */
	public String setRegion(String value) throws Exception {
		m_region = value;
		return value;
	}

	/**
	 * Gets the postal code of the employee.
	 */
	@XmlElement(name = "PostalCode")
	public String getPostalCode() throws Exception {
		return m_postalCode;
	}

	/**
	 * Sets the postal code of the employee.
	 */
	public String setPostalCode(String value) throws Exception {
		m_postalCode = value;
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
	 * Gets the contact number of the employee.
	 */
	@XmlElement(name = "HomePhone")
	public String getHomePhone() throws Exception {
		return m_homePhone;
	}

	/**
	 * Sets the contact number of the employee.
	 */
	public String setHomePhone(String value) throws Exception {
		m_homePhone = value;
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
	 * Gets the photo of the employee.
	 */
	@XmlElement(name = "Photo")
	public String getPhoto() throws Exception {
		return m_photo;
	}

	/**
	 * Sets the photo of the employee.
	 */
	public String setPhoto(String value) throws Exception {
		m_photo = value;
		return value;
	}

	/**
	 * Gets the notes.
	 */
	@XmlElement(name = "Notes")
	public String getNotes() throws Exception {
		return m_notes;
	}

	/**
	 * Sets the notes.
	 */
	public String setNotes(String value) throws Exception {
		m_notes = value;
		return value;
	}

	/**
	 * Gets the reporter of the employee.
	 */
	@XmlElement(name = "ReportsTo")
	public String getReportsTo() throws Exception {
		return m_reportsTo;
	}

	/**
	 * Sets the reporter of the employee.
	 */
	public String setReportsTo(String value) throws Exception {
		m_reportsTo = value;
		return value;
	}
}
