package salesinvoice;

import java.util.List;
import javax.xml.bind.annotation.*;

@XmlRootElement(name = "TestOrderDetails")
public class TestOrderDetails {
	private List<TestOrderDetail> TestOrderDetail;

	/**
	 * Gets the list of order details.
	 */
	@XmlElement(name = "TestOrderDetail")
	public List<TestOrderDetail> getTestOrderDetails() {
		return TestOrderDetail;
	}

	/**
	 * Sets the list of order details.
	 */
	public void setTestOrderDetails(List<TestOrderDetail> testOrderDetail) {
		this.TestOrderDetail = testOrderDetail;
	}
}
