package salesinvoice;

import java.util.List;

import javax.xml.bind.annotation.*;

@XmlRootElement(name = "TestOrderTotals")
public class TestOrderTotals {
	private List<TestOrderTotal> TestOrderTotal;

	/**
	 * Gets the list of total of the order.
	 */
	@XmlElement(name = "TestOrderTotal")
	public List<TestOrderTotal> getTestOrderTotals() {
		return TestOrderTotal;
	}

	/**
	 * Sets the list of total of the order.
	 */
	public void setTestOrderTotals(List<TestOrderTotal> testOrderTotal) {
		this.TestOrderTotal = testOrderTotal;
	}
}
