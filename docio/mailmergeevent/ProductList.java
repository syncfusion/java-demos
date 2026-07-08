package mailmergeevent;

import java.util.List;
import javax.xml.bind.annotation.*;

@XmlRootElement(name = "ProductList")
public class ProductList {
	private List<Products> Products;

	/**
	 * Gets the list of product.
	 */
	@XmlElement(name = "Products")
	public List<Products> getProducts() {
		return Products;
	}

	/**
	 * Sets the list of product.
	 */
	public void setProducts(List<Products> products) {
		this.Products = products;
	}
}
