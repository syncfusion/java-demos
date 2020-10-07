package mailmergeevent;

public class ProductDetail {
	private String m_sNO;
	private String m_productName;
	private String m_productImage;

	public String getSNO() throws Exception {
		return m_sNO;
	}

	public String setSNO(String value) throws Exception {
		m_sNO = value;
		return value;
	}

	public String getProductName() throws Exception {
		return m_productName;
	}

	public String setProductName(String value) throws Exception {
		m_productName = value;
		return value;
	}

	public String getProductImage() throws Exception {
		return m_productImage;
	}

	public String setProductImage(String value) throws Exception {
		m_productImage = value;
		return value;
	}

	public ProductDetail(String sNO, String productName, String productImage) throws Exception {
		m_sNO = sNO;
		m_productName = productName;
		m_productImage = productImage;
	}

	public ProductDetail() throws Exception {
	}
}
