package tablestyles;

import java.util.List;
import javax.xml.bind.annotation.*;

@XmlRootElement(name = "SuppliersList")
public class SuppliersList {
	private List<Suppliers> Suppliers;

	/**
	 * Gets the list of supplier details.
	 */
	@XmlElement(name = "Suppliers")
	public List<Suppliers> getSuppliers() {
		return Suppliers;
	}

	/**
	 * Sets the list of supplier details.
	 */
	public void setSuppliers(List<Suppliers> supplier) {
		this.Suppliers = supplier;
	}
}
