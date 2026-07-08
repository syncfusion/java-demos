package salesinvoice;

import javax.xml.bind.annotation.XmlElement;

public class TestOrder {
	private String m_orderID;
	private String m_shipName;
	private String m_shipAddress;
	private String m_shipCity;
	private String m_shipPostalCode;
	private String m_shipCountry;
	private String m_customerID;
	private String m_address;
	private String m_postalCode;
	private String m_city;
	private String m_country;
	private String m_salesPerson;
	private String m_customersCompanyName;
	private String m_orderDate;
	private String m_requiredDate;
	private String m_shippedDate;
	private String m_shippersCompanyName;

	/**
	 * Gets the ship name of order.
	 */
	@XmlElement(name = "ShipName")
	public String getShipName() throws Exception {
		return m_shipName;
	}

	/**
	 * Sets the ship name of order.
	 */
	public String setShipName(String value) throws Exception {
		m_shipName = value;
		return value;
	}

	/**
	 * Gets the shipping address of order.
	 */
	@XmlElement(name = "ShipAddress")
	public String getShipAddress() throws Exception {
		return m_shipAddress;
	}

	/**
	 * Sets the shipping address of order.
	 */
	public String setShipAddress(String value) throws Exception {
		m_shipAddress = value;
		return value;
	}

	/**
	 * Gets the shipping city of order.
	 */
	@XmlElement(name = "ShipCity")
	public String getShipCity() throws Exception {
		return m_shipCity;
	}

	/**
	 * Sets the shipping city of order.
	 */
	public String setShipCity(String value) throws Exception {
		m_shipCity = value;
		return value;
	}

	/**
	 * Gets the shipping postal code of order.
	 */
	@XmlElement(name = "ShipPostalCode")
	public String getShipPostalCode() throws Exception {
		return m_shipPostalCode;
	}

	/**
	 * Sets the shipping postal code of order.
	 */
	public String setShipPostalCode(String value) throws Exception {
		m_shipPostalCode = value;
		return value;
	}

	/**
	 * Gets the postal code of order.
	 */
	@XmlElement(name = "PostalCode")
	public String getPostalCode() throws Exception {
		return m_postalCode;
	}

	/**
	 * Sets the postal code of order.
	 */
	public String setPostalCode(String value) throws Exception {
		m_postalCode = value;
		return value;
	}

	/**
	 * Gets the shipping country of order.
	 */
	@XmlElement(name = "ShipCountry")
	public String getShipCountry() throws Exception {
		return m_shipCountry;
	}

	/**
	 * Sets the shipping country of order.
	 */
	public String setShipCountry(String value) throws Exception {
		m_shipCountry = value;
		return value;
	}

	/**
	 * Sets the ID of the customer.
	 */
	@XmlElement(name = "CustomerID")
	public String getCustomerID() throws Exception {
		return m_customerID;
	}

	/**
	 * Sets the shipping country of order.
	 */
	public String setCustomerID(String value) throws Exception {
		m_customerID = value;
		return value;
	}

	/**
	 * Gets the company name of the customer.
	 */
	@XmlElement(name = "Customers.CompanyName")
	public String getCustomers_CompanyName() throws Exception {
		return m_customersCompanyName;
	}

	/**
	 * Sets the company name of the customer.
	 */
	public String setCustomers_CompanyName(String value) throws Exception {
		m_customersCompanyName = value;
		return value;
	}

	/**
	 * Gets the address of the customer.
	 */
	@XmlElement(name = "Address")
	public String getAddress() throws Exception {
		return m_address;
	}

	/**
	 * Sets the address of the customer.
	 */
	public String setAddress(String value) throws Exception {
		m_address = value;
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
	 * Gets the sales person details.
	 */
	@XmlElement(name = "Salesperson")
	public String getSalesperson() throws Exception {
		return m_salesPerson;
	}

	/**
	 * Sets the sales person details.
	 */
	public String setSalesperson(String value) throws Exception {
		m_salesPerson = value;
		return value;
	}

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
	 * Gets the date of the order.
	 */
	@XmlElement(name = "OrderDate")
	public String getOrderDate() throws Exception {
		return m_orderDate;
	}

	/**
	 * Sets the date of the order.
	 */
	public String setOrderDate(String value) throws Exception {
		int index = value.indexOf("T");
		m_orderDate = value.substring(0, index);
		return value;
	}

	/**
	 * Gets the required date of the order.
	 */
	@XmlElement(name = "RequiredDate")
	public String getRequiredDate() throws Exception {
		return m_requiredDate;
	}

	/**
	 * Sets the required date of the order.
	 */
	public String setRequiredDate(String value) throws Exception {
		int index = value.indexOf("T");
		m_requiredDate = value.substring(0, index);
		return value;
	}

	/**
	 * Gets the shipped date of the order.
	 */
	@XmlElement(name = "ShippedDate")
	public String getShippedDate() throws Exception {
		return m_shippedDate;
	}

	/**
	 * Sets the shipped date of the order.
	 */
	public String setShippedDate(String value) throws Exception {
		int index = value.indexOf("T");
		m_shippedDate = value.substring(0, index);
		return value;
	}

	/**
	 * Gets the company name of the shipper.
	 */
	@XmlElement(name = "Shippers.CompanyName")
	public String getShippers_CompanyName() throws Exception {
		return m_shippersCompanyName;
	}

	/**
	 * Sets the company name of the shipper.
	 */
	public String setShippers_CompanyName(String value) throws Exception {
		m_shippersCompanyName = value;
		return value;
	}

}
