package mailmergeevent;

import java.util.List;
import javax.xml.bind.annotation.*;

@XmlRootElement(name = "Products")
public class ProductDetail {
	private List<Product_PriceList> PriceList;

	/**
	 * Gets the list of product price.
	 */
	@XmlElement(name = "Product_PriceList")
	public List<Product_PriceList> getPriceList() {
		return PriceList;
	}

	/**
	 * Sets the list of product price.
	 */
	public void setPriceList(List<Product_PriceList> priceList) {
		this.PriceList = priceList;
	}

}
