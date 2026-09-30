package updatefields;

import java.util.List;
import javax.xml.bind.annotation.*;

@XmlRootElement(name = "StockMarket")
public class StockMarket {
	private List<StockDetails> StockDetails;

	/**
	 * Gets the list of stock details.
	 */
	@XmlElement(name = "StockDetails")
	public List<StockDetails> getStockDetails() {
		return StockDetails;
	}

	/**
	 * Sets the list of stock details.
	 */
	public void setStockDetails(List<StockDetails> listStockDetails) {
		this.StockDetails = listStockDetails;
	}
}
