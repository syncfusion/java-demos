package updatefields;

import javax.xml.bind.annotation.XmlElement;

public class StockDetails {
	private String m_tradeNo;
	private String m_companyName;
	private String m_costPrice;
	private String m_sharesCount;
	private String m_salesPrice;

	/**
	 * Gets the trade number.
	 */
	@XmlElement(name = "TradeNo")
	public String getTradeNo() throws Exception {
		return m_tradeNo;
	}

	/**
	 * Sets the trade number.
	 */
	public String setTradeNo(String value) throws Exception {
		m_tradeNo = value;
		return value;
	}

	/**
	 * Gets the name of the company.
	 */
	@XmlElement(name = "CompanyName")
	public String getCompanyName() throws Exception {
		return m_companyName;
	}

	/**
	 * Sets the name of the company.
	 */
	public String setCompanyName(String value) throws Exception {
		m_companyName = value;
		return value;
	}

	/**
	 * Gets the cost price.
	 */
	@XmlElement(name = "CostPrice")
	public String getCostPrice() throws Exception {
		return m_costPrice;
	}

	/**
	 * Sets the cost price.
	 */
	public String setCostPrice(String value) throws Exception {
		m_costPrice = value;
		return value;
	}

	/**
	 * Gets the total count of the share.
	 */
	@XmlElement(name = "SharesCount")
	public String getSharesCount() throws Exception {
		return m_sharesCount;
	}

	/**
	 * Sets the total count of the share.
	 */
	public String setSharesCount(String value) throws Exception {
		m_sharesCount = value;
		return value;
	}

	/**
	 * Gets the sales price.
	 */
	@XmlElement(name = "SalesPrice")
	public String getSalesPrice() throws Exception {
		return m_salesPrice;
	}

	/**
	 * Sets the sales price.
	 */
	public String setSalesPrice(String value) throws Exception {
		m_salesPrice = value;
		return value;
	}

}
