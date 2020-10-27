package salesinvoice;

import java.text.SimpleDateFormat;
import java.util.Date;  

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

	public String getShipName() throws Exception {
		return m_shipName;
	}

	public String setShipName(String value) throws Exception {
		m_shipName = value;
		return value;
	}

	public String getShipAddress() throws Exception {
		return m_shipAddress;
	}

	public String setShipAddress(String value) throws Exception {
		m_shipAddress = value;
		return value;
	}

	public String getShipCity() throws Exception {
		return m_shipCity;
	}

	public String setShipCity(String value) throws Exception {
		m_shipCity = value;
		return value;
	}

	public String getShipPostalCode() throws Exception {
		return m_shipPostalCode;
	}

	public String setShipPostalCode(String value) throws Exception {
		m_shipPostalCode = value;
		return value;
	}

	public String getPostalCode() throws Exception {
		return m_postalCode;
	}

	public String setPostalCode(String value) throws Exception {
		m_postalCode = value;
		return value;
	}

	public String getShipCountry() throws Exception {
		return m_shipCountry;
	}

	public String setShipCountry(String value) throws Exception {
		m_shipCountry = value;
		return value;
	}

	public String getCustomerID() throws Exception {
		return m_customerID;
	}

	public String setCustomerID(String value) throws Exception {
		m_customerID = value;
		return value;
	}

	public String getCustomers_CompanyName() throws Exception {
		return m_customersCompanyName;
	}

	public String setCustomers_CompanyName(String value) throws Exception {
		m_customersCompanyName = value;
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

	public String getSalesperson() throws Exception {
		return m_salesPerson;
	}

	public String setSalesperson(String value) throws Exception {
		m_salesPerson = value;
		return value;
	}

	public String getOrderID() throws Exception {
		return m_orderID;
	}

	public String setOrderID(String value) throws Exception {
		m_orderID = value;
		return value;
	}

	public String getOrderDate() throws Exception {
		return m_orderDate;
	}

	public String setOrderDate(String value) throws Exception {	
		int index = value.indexOf("T");		
		m_orderDate = value.substring(0, index);	
		return value;
	}

	public String getRequiredDate() throws Exception {
		return m_requiredDate;
	}

	public String setRequiredDate(String value) throws Exception {
		int index = value.indexOf("T");		
		m_requiredDate = value.substring(0, index);		
		return value;
	}

	public String getShippedDate() throws Exception {
		return m_shippedDate;
	}

	public String setShippedDate(String value) throws Exception {
		int index = value.indexOf("T");		
		m_shippedDate = value.substring(0, index);		
		return value;
	}

	public String getShippers_CompanyName() throws Exception {
		return m_shippersCompanyName;
	}

	public String setShippers_CompanyName(String value) throws Exception {
		m_shippersCompanyName = value;
		return value;
	}

	public TestOrder() throws Exception {
	}
}
