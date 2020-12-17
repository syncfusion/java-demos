package salesinvoice;

public class TestOrderTotal {
	private String m_orderID;
	private String m_subTotal;
	private String m_freight;
	private String m_total;

	public String getOrderID() throws Exception {
		return m_orderID;
	}

	public String setOrderID(String value) throws Exception {
		m_orderID = value;
		return value;
	}

	public String getSubtotal() throws Exception {
		return m_subTotal;
	}

	public String setSubtotal(String value) throws Exception {
		m_subTotal = value;
		return value;
	}

	public String getFreight() throws Exception {
		return m_freight;
	}

	public String setFreight(String value) throws Exception {
		m_freight = value;
		return value;
	}

	public String getTotal() throws Exception {
		return m_total;
	}

	public String setTotal(String value) throws Exception {
		m_total = value;
		return value;
	}

	public TestOrderTotal() throws Exception {
	}
}
