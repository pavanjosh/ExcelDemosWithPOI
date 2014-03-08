package model;

import java.util.ArrayList;
import java.util.List;

public class Model {

	private List<String> missingNames = new ArrayList<String>();
	private ExcelSheet sheet = null;
	private ExcelSheet missedShiftsheet = null;
	private List<Employee> employeeList = null;
	private List<String> diffInWorkingHoursNames = new ArrayList<String>();
	private List<Timesheet> timeSheetList = null;
	private List<String> peopleNotCompared = new ArrayList<String>();
	private List<Invoice> invoiceList = null;
	
	private List<Invoice> notComparedInvoiceList = null;
	
	public List<String> getPeopleNotCompared() {
		return peopleNotCompared;
	}
	public void setPeopleNotCompared(List<String> peopleNotCompared) {
		this.peopleNotCompared = peopleNotCompared;
	}
	public ExcelSheet getSheet() {
		return sheet;
	}
	public void setSheet(ExcelSheet sheet) {
		this.sheet = sheet;
	}
	public List<Employee> getEmployeeList() {
		return employeeList;
	}
	public void setEmployeeList(List<Employee> employeeList) {
		this.employeeList = employeeList;
	}
	public void addMissingNames(String name) {
		missingNames.add(name);
		
	}
	public List<String> getMissingNames() {
		return missingNames;
	}
	public void setMissingNames(List<String> missingNames) {
		this.missingNames = missingNames;
	}
	public List<Timesheet> getTimeSheetList() {
		return timeSheetList;
	}
	public void setTimeSheetList(List<Timesheet> timeSheetList) {
		this.timeSheetList = timeSheetList;
	}
	public List<String> getDiffInWorkingHoursNames() {
		return diffInWorkingHoursNames;
	}
	public void setDiffInWorkingHoursNames(List<String> diffInWorkingHoursNames) {
		this.diffInWorkingHoursNames = diffInWorkingHoursNames;
	}
	public boolean isEmployeeToBeCompared(String name) {
		
		if(employeeList != null)
		{
			for (Employee emploee : employeeList) {

				if(emploee.getName().equalsIgnoreCase(name))
				{
					if(emploee.isToBeComapred())
					{
						return true;
					}
					else
					{
						return false;
					}
				}

			}
		}
		return false;
	}
	public ExcelSheet getMissedShiftsheet() {
		return missedShiftsheet;
	}
	public void setMissedShiftsheet(ExcelSheet missedShiftsheet) {
		this.missedShiftsheet = missedShiftsheet;
	}
	public Employee getEmployee(String name) {
		for (Employee employee : employeeList) {
			if(employee.getName().equalsIgnoreCase(name))
			{
				return employee;
			}
		}
		return null;
	}
	public List<Invoice> getInvoiceList() {
		return invoiceList;
	}
	public void setInvoiceList(List<Invoice> invoiceList) {
		this.invoiceList = invoiceList;
	}
	public List<Invoice> getNotComparedInvoiceList() {
		return notComparedInvoiceList;
	}
	public void setNotComparedInvoiceList(List<Invoice> notComparedInvoiceList) {
		this.notComparedInvoiceList = notComparedInvoiceList;
	}
	
}
