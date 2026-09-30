package salesinvoice;

import javax.xml.bind.annotation.XmlElement;

public class TestOrderTotal {
	private String m_orderID;
	private String m_subTotal;
	private String m_freight;
	private String m_total;

	/**
	 * Gets the ID of the order.
	 */
	@XmlElement(name = "OrderID")
	public String getOrderID() throws Exception {
		return m_orderID;
	}

	/**
	 * Sets the ID of the order.
	 */
	public String setOrderID(String value) throws Exception {
		m_orderID = value;
		return value;
	}

	/**
	 * Gets the sub total of the order.
	 */
	@XmlElement(name = "Subtotal")
	public String getSubtotal() throws Exception {
		return m_subTotal;
	}

	/**
	 * Sets the sub total of the order.
	 */
	public String setSubtotal(String value) throws Exception {
		m_subTotal = value;
		return value;
	}

	/**
	 * Gets the freight of the order.
	 */
	@XmlElement(name = "Freight")
	public String getFreight() throws Exception {
		return m_freight;
	}

	/**
	 * Sets the freight of the order.
	 */
	public String setFreight(String value) throws Exception {
		m_freight = value;
		return value;
	}

	/**
	 * Gets the total of the order.
	 */
	@XmlElement(name = "Total")
	public String getTotal() throws Exception {
		return m_total;
	}

	/**
	 * Sets the total of the order.
	 */
	public String setTotal(String value) throws Exception {
		m_total = value;
		return value;
	}

}
