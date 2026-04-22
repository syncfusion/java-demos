package nestedmailmerge;

import java.util.List;

import javax.xml.bind.annotation.*;

import com.syncfusion.javahelper.system.collections.generic.ListSupport;

@XmlRootElement(name = "EmployeesList")
public class Employees {

	private List<EmployeeDetails> Employee;

	/**
	 * Gets the list of employee details.
	 */
	@XmlElement(name = "Employees")
	public List<EmployeeDetails> getEmployees() throws Exception {
		return Employee;
	}

	/**
	 * Sets the list of employee details.
	 */
	public void setEmployees(List<EmployeeDetails> employee) {
		this.Employee = employee;
	}

	/**
	 * Gets the list of customer details.
	 */
	public ListSupport<CustomerDetails> getCustomers() throws Exception {
		ListSupport<CustomerDetails> customersList = new ListSupport<CustomerDetails>(CustomerDetails.class);
		for (EmployeeDetails employee : Employee) {
			customersList.addRange(employee.getCustomerDetails());
		}
		return customersList;
	}

	/**
	 * Gets the list of order details.
	 */
	public ListSupport<OrderDetails> getOrders() throws Exception {
		ListSupport<OrderDetails> customersList = new ListSupport<OrderDetails>(OrderDetails.class);
		for (EmployeeDetails employee : Employee) {
			for (CustomerDetails customer : employee.getCustomerDetails()) {
				customersList.addRange(customer.getOrders());
			}
		}
		return customersList;
	}

}
