package mailmergeevent;

import javax.xml.bind.annotation.XmlElement;

public class Products {
	private String m_sNO;
	private String m_productName;
	private String m_productImage;

	/**
	 * Gets the serial number of the product.
	 */
	@XmlElement(name = "SNO")
	public String getSNO() throws Exception {
		return m_sNO;
	}

	/**
	 * Sets the serial number of the product.
	 */
	public String setSNO(String value) throws Exception {
		m_sNO = value;
		return value;
	}

	/**
	 * Gets the name of the product.
	 */
	@XmlElement(name = "ProductName")
	public String getProductName() throws Exception {
		return m_productName;
	}

	/**
	 * Sets the name of the product.
	 */
	public String setProductName(String value) throws Exception {
		m_productName = value;
		return value;
	}

	/**
	 * Gets the image of the product.
	 */
	@XmlElement(name = "ProductImage")
	public String getProductImage() throws Exception {
		return m_productImage;
	}

	/**
	 * Sets the image of the product.
	 */
	public String setProductImage(String value) throws Exception {
		m_productImage = value;
		return value;
	}
}
