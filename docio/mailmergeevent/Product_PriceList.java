package mailmergeevent;

public class Product_PriceList {
	private String m_productName;
	private String m_price;

	public String getProductName() throws Exception {
		return m_productName;
	}

	public String setProductName(String value) throws Exception {
		m_productName = value;
		return value;
	}

	public String getPrice() throws Exception {
		return m_price;
	}

	public String setPrice(String value) throws Exception {
		m_price = value;
		return value;
	}

	public Product_PriceList(String productName, String price) throws Exception {
		m_productName = productName;
		m_price = price;
	}

	public Product_PriceList() throws Exception {
	}
}
