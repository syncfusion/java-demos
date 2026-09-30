package salesinvoice;

import java.util.List;
import javax.xml.bind.annotation.*;

@XmlRootElement(name = "TestOrders")
public class TestOrders {
	private List<TestOrder> TestOrder;

	/**
	 * Gets the list of order.
	 */
	@XmlElement(name = "TestOrder")
	public List<TestOrder> getTestOrder() {
		return TestOrder;
	}

	/**
	 * Sets the list of order.
	 */
	public void setTestOrder(List<TestOrder> testOrder) {
		this.TestOrder = testOrder;
	}

}
