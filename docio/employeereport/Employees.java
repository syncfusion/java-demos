package employeereport;

import java.util.List;
import javax.xml.bind.annotation.*;

@XmlRootElement(name = "Employees")
public class Employees {
	private List<Employee> Employee;

	/**
	 * Gets the list of the employee.
	 */
	@XmlElement(name = "Employee")
	public List<Employee> getEmployees() {
		return Employee;
	}

	/**
	 * Sets the list of the employee.
	 */
	public void setEmployees(List<Employee> employee) {
		this.Employee = employee;
	}
}
