package businesslogic;

import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.TreeMap;

import model.CellStyles;
import model.Constants;
import model.Employee;
import model.ExcelCellData;
import model.ExcelSheet;
import model.ExcelSheet.Person;
import model.Invoice;
import model.Model;
import model.Timesheet;

public class Operations {
	public Model execute(Model model){
		
		calculateTotalWorkedHours(model);
		//calcultaePublicHolidayRate();
		calculateDiffTimeValues(model);
		listMissingNamesFromMasterSheet(model);
		ExcelSheet missedShiftsheet = model.getMissedShiftsheet();
		if(missedShiftsheet != null)
		{
			adjustMissedShiftRates(model);
		}
		generateNotComparedInvoices(model);
		return model;
	}
 
	private void generateNotComparedInvoices(Model model) {
		List<Employee> employeeList = model.getEmployeeList();
		List<Invoice> notComparedInvoiceList = new ArrayList<Invoice>();
		ExcelSheet masterSheet = model.getSheet();
		for (Employee employee : employeeList) {
			if(!employee.isToBeComapred())
			{
				String name = employee.getName();
				Person person = masterSheet.getPerson(name);
				if(person != null)
				{
					Double totalWorkingHours = person.getTotalWorkingHours();
					Double generalRate = employee.getGeneralRate();
					Double invoiceRate = totalWorkingHours * generalRate;
					Invoice invoice = new Invoice();
					invoice.setFinalRate(invoiceRate);
					invoice.setFinalTotalWorkedHours(totalWorkingHours);
					invoice.setFullName(name);
					notComparedInvoiceList.add(invoice);
				}
			}
		}
		model.setNotComparedInvoiceList(notComparedInvoiceList);
		
		
	}

	private void adjustMissedShiftRates(Model model) {
		
		ExcelSheet missedShiftsheet = model.getMissedShiftsheet();
		ExcelSheet masterSheet = model.getSheet();
		List<Invoice> invoiceList = new ArrayList<Invoice>();
		
		List<Person> missedShiftPersonList = missedShiftsheet.getPersonList();
		for (Person missedPerson : missedShiftPersonList) {
			String name = missedPerson.getName();
			Person person = masterSheet.getPerson(name);
			if(person != null)
			{
				Double totalWorkingHours = person.getTotalWorkingHours();
				Double missedPersonTotalWorkingHours = missedPerson.getTotalWorkingHours();
				Double finalHours = totalWorkingHours+missedPersonTotalWorkingHours;
				Employee employee = model.getEmployee(name);
				if(employee != null)
				{
					Double finalRate = finalHours*employee.getGeneralRate();

					Invoice invoice = new Invoice();
					invoice.setFullName(name);
					invoice.setFirstName(employee.getFirstName());
					invoice.setLastName(employee.getLasttName());

					invoice.setFinalRate(finalRate);
					invoice.setFinalTotalWorkedHours(finalHours);
					
					invoiceList.add(invoice);
				}
			}


		}
		model.setInvoiceList(invoiceList);
	}

	private void calculateDiffTimeValues(Model model) {
		List<Timesheet> timeSheetList = model.getTimeSheetList();
		if(timeSheetList != null)
		{
			ExcelSheet sheet = model.getSheet();
			double diffHours = 0.0;
			List<String> diffInWorkingHours = model.getDiffInWorkingHoursNames();
			for (Timesheet timesheet : timeSheetList) {
				String name = timesheet.getName();
				Double totalWorkingHours = sheet.getTotalWorkingHours(name);
				Double weeklyWorkedHours = timesheet.getWeeklyWorkedHours();

				if(totalWorkingHours != -1.0)
				//if(totalWorkingHours != -1.0 && weeklyWorkedHours != -1.0)
				{
					int weeklyHoursCompare = weeklyWorkedHours.compareTo(-1.0);
					if(weeklyHoursCompare == 0)
					{
						weeklyWorkedHours = 0.0;
					}
					int comparison = weeklyWorkedHours.compareTo(-1000.00);
					if(comparison == 0)
					{
						sheet.setWorkingHoursMatch(true,name);
						model.getPeopleNotCompared().add(name);
					}
					else{
						int diff = totalWorkingHours.compareTo(weeklyWorkedHours);
						if(diff != 0)
						{
							diffInWorkingHours.add(name);
							int quantityRateColumnNum = sheet.getColNum(Constants.QUANTITY_COLUMN_NAME);
							int startRow = sheet.getStartRow(name);
							quantityRateColumnNum++;
							ExcelCellData cell = new ExcelCellData();
							cell.setColNum(quantityRateColumnNum);
							cell.setRowNum(startRow);
							//if(diff >=0)
							{
								diffHours = totalWorkingHours - weeklyWorkedHours;
							}

							cell.setValue(diffHours);
							sheet.setTotalDiffHours(name,diffHours);
							//sheet.addCell(name,cell);
						}
						else
						{
							sheet.setWorkingHoursMatch(true,name);
						}
					}
				}
			}
		}
		

	}

	private void calculateTotalWorkedHours(Model model) {


		ExcelSheet sheet = model.getSheet();

		List<Person> personList = sheet.getPersonList();
		for (Person person : personList) {
			String name = person.getName();

			List<Integer> rowList =  new ArrayList<Integer>();
			rowList.add(person.getStartRow());
			rowList.add(person.getEndRow());
			
			
			//List<Integer> rowList = sheet.getStartAndEndRowNums(name);
			Employee employee = getEmployee(name,model);
			
			
			//int endRow = sheet.getEndRow(name);
			
			int endRow =  rowList.get(1);
			
			int quantityRateColumnNum = sheet.getColNum(Constants.QUANTITY_COLUMN_NAME);
			ExcelCellData cell = new ExcelCellData();
			cell.setColNum(quantityRateColumnNum);
			cell.setRowNum(endRow);
			
			cell.setSpecialCell(true);
			char c = getAlphabet(quantityRateColumnNum);
			
			String formula = "SUM"+"("+c + (rowList.get(0)+1) +":"+c + (endRow) +")";
			cell.setFormula(formula);
			cell.setSetStyle(true);
			cell.setStyle(CellStyles.TOTAL_HOURS_STYLE);
			
			
			sheet.addCell(name,cell);

			
			ExcelCellData payRateCell = new ExcelCellData();
			int payRateColNum = quantityRateColumnNum+1;
			
			payRateCell.setColNum(payRateColNum);
			payRateCell.setRowNum(endRow);
			if(employee != null)
			{
				payRateCell.setValue(employee.getGeneralRate());
			}
			else
			{
				payRateCell.setValue(0);
			}
			
			payRateCell.setSetStyle(true);
			sheet.addCell(name,payRateCell);
			
			
			ExcelCellData totalSumRateCell = new ExcelCellData();
			int totalSumColumnNumber = quantityRateColumnNum + 2;
			
			totalSumRateCell.setColNum(totalSumColumnNumber);
			totalSumRateCell.setRowNum(endRow);
			
			
			char rateColumn = getAlphabet(payRateColNum);
			
			formula = "PRODUCT"+"("+c + (endRow+1) +":"+rateColumn + (endRow+1) +")";
			totalSumRateCell.setFormula(formula);
			totalSumRateCell.setSpecialCell(true);
			totalSumRateCell.setSetStyle(true);
			sheet.addCell(name,totalSumRateCell);
		}
		
		
		
		
//		for (Employee employee : employeeList) {
//
//			String name = employee.getName();
//
//			if(!sheet.isNamePresent(name))
//			{
//				//model.addMissingNames(name);
//			}
//			else
//			{
//				List<Integer> rowList = sheet.getStartAndEndRowNums(name);
//				int endRow = sheet.getEndRow(name);
//				int quantityRateColumnNum = sheet.getColNum(Constants.QUANTITY_COLUMN_NAME);
//				ExcelCellData cell = new ExcelCellData();
//				cell.setColNum(quantityRateColumnNum);
//				cell.setRowNum(endRow);
//				
//				cell.setSpecialCell(true);
//				char c = getAlphabet(quantityRateColumnNum);
//				
//				String formula = "SUM"+"("+c + (rowList.get(0)+1) +":"+c + (endRow) +")";
//				cell.setFormula(formula);
//				cell.setSetStyle(true);
//				cell.setStyle(CellStyles.TOTAL_HOURS_STYLE);
//				
//				
//				sheet.addCell(name,cell);
//
//				
//				ExcelCellData payRateCell = new ExcelCellData();
//				int payRateColNum = quantityRateColumnNum+1;
//				
//				payRateCell.setColNum(payRateColNum);
//				payRateCell.setRowNum(endRow);
//				payRateCell.setValue(employee.getGeneralRate());
//				payRateCell.setSetStyle(true);
//				sheet.addCell(name,payRateCell);
//				
//				
//				ExcelCellData totalSumRateCell = new ExcelCellData();
//				int totalSumColumnNumber = quantityRateColumnNum + 2;
//				
//				totalSumRateCell.setColNum(totalSumColumnNumber);
//				totalSumRateCell.setRowNum(endRow);
//				
//				
//				char rateColumn = getAlphabet(payRateColNum);
//				
//				formula = "PRODUCT"+"("+c + (endRow+1) +":"+rateColumn + (endRow+1) +")";
//				totalSumRateCell.setFormula(formula);
//				totalSumRateCell.setSpecialCell(true);
//				totalSumRateCell.setSetStyle(true);
//				sheet.addCell(name,totalSumRateCell);
//
//				
//			}
//		}
	}

	private Employee getEmployee(String name , Model model) {
	
		List<Employee> employeeList = model.getEmployeeList();
		for (Employee employee : employeeList) {
			if(employee.getName().equalsIgnoreCase(name))
			{
				return employee;
			}
		}
		return null;
	}

	public void listMissingNamesFromMasterSheet(Model model)
	{
		boolean namePresent = false;
		List<Employee> employeeList = model.getEmployeeList();
		List<Person> personList = model.getSheet().getPersonList();
		for (Person person : personList) {
			String name = person.getName();
			namePresent = false;
			for (Employee employee : employeeList) {
				int indexOf = employeeList.indexOf(employee);
				if(employee.getName()!=null)
				{
					if(employee.getName().equalsIgnoreCase(name))
					{
						namePresent = true;
						break;
					}
				}
				
			}
			if(namePresent == false)
			{
				model.getMissingNames().add(name);
			}
		}
		
	}
	
	private char getAlphabet(int colNum) {
		Map<Integer,Character> alphabetMap = new TreeMap<Integer, Character>();
		alphabetMap.put(0, 'A');
		alphabetMap.put(1, 'B');
		alphabetMap.put(2, 'C');
		alphabetMap.put(3, 'D');
		alphabetMap.put(4, 'E');
		alphabetMap.put(5, 'F');
		alphabetMap.put(6, 'G');
		alphabetMap.put(7, 'H');
		alphabetMap.put(8, 'I');
		alphabetMap.put(9, 'J');
		alphabetMap.put(10, 'K');
		alphabetMap.put(11, 'L');
		alphabetMap.put(12, 'M');
		alphabetMap.put(13, 'N');
		alphabetMap.put(14, 'O');
		alphabetMap.put(15, 'P');
		alphabetMap.put(16, 'Q');
		alphabetMap.put(17, 'R');
		alphabetMap.put(18, 'S');
		alphabetMap.put(19, 'T');
		alphabetMap.put(20, 'U');
		alphabetMap.put(21, 'V');
		alphabetMap.put(22, 'W');
		alphabetMap.put(23, 'X');
		alphabetMap.put(24, 'Y');
		alphabetMap.put(25, 'Z');
	
		return alphabetMap.get(colNum);
		
	}

}
