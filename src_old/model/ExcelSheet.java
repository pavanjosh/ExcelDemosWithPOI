package model;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import javax.swing.text.StyleContext;

import model.ExcelSheet.Person;

public class ExcelSheet {
	
	//String currentName = null;
	//Person personObj = new Person();
	
	Person currentPerson = null;
	
	int currentRowNumber = 0;

	boolean personFound = false;
	
	private int endColumnNumber = 0;
	
	
	//boolean personObjFound = false;
	//boolean firstEntry = true;
	//int currentRowNum = 0;
	
	private Map<String,ExcelCellData> sheetData = new HashMap<String,ExcelCellData>();

	private List<ExcelCellData> headerCells = new ArrayList<ExcelCellData>();
	
	//private Map<Person, List<ExcelCellData>> personToCellDataMap = new HashMap<Person, List<ExcelCellData>>();
	
	private List<Person> personList = new ArrayList<Person>();
	
	
//	public Map<Person, List<ExcelCellData>> getPersonToCellDataMap() {
//		return personToCellDataMap;
//	}
//
//	public void setPersonToCellDataMap(
//			Map<Person, List<ExcelCellData>> personToCellDataMap) {
//		this.personToCellDataMap = personToCellDataMap;
//	}

	public List<Person> getPersonList() {
		return personList;
	}

	public void setPersonList(List<Person> personList) {
		this.personList = personList;
	}
	public List<ExcelCellData> getHeaderCells() {
		return headerCells;
	}

	public void setHeaderCells(List<ExcelCellData> headerCells) {
		this.headerCells = headerCells;
	}

	public Map<String, ExcelCellData> getSheetData() {
		return sheetData;
	}

	public void setSheetData(Map<String, ExcelCellData> sheetData) {
		this.sheetData = sheetData;
	}

	public void populateSheetData(String name, List<ExcelCellData> cellList)
	{
		
		Person tempPerson = null;
		personFound = false;
		for (Person person : personList) {
			if(person.getName().equalsIgnoreCase(name))
			{
				
				personFound = true;
				tempPerson = person;
				break;
				
			}
		}
		if(personFound == false)
		{
			Person newPerson = new Person();
			newPerson.setName(name);

			if(currentPerson == null)
			{
				currentRowNumber = cellList.get(0).getRowNum();
				
			}
			else
			{
				//currentRowNumber++;
				currentPerson.setEndRow(currentRowNumber);
				currentRowNumber++;

			}
			
			newPerson.setStartRow(currentRowNumber);
			newPerson.setEndRow(currentRowNumber);
			newPerson.addCellDataList(cellList,currentRowNumber);
			currentPerson = newPerson;
			this.personList.add(newPerson);
		}
		else
		{
			int indexOfPerson = personList.indexOf(tempPerson);
			Person person = personList.get(indexOfPerson);
			person.addCellDataList(cellList,currentRowNumber);
			person.setEndRow(currentRowNumber);
		}
		currentRowNumber++;



//		if(firstEntry)
//		{
//			firstEntry = false;
//
//			Person person = new Person();
//			person.setName(name);
//			currentName = name;
//			List<ExcelCellData> personToCellDataList = new ArrayList<ExcelCellData>();
//
//			currentRowNum = cellList.get(0).getRowNum();
//			person.setStartRow(currentRowNum);
//			person.setEndRow(currentRowNum);
//
//			personToCellDataList.addAll(cellList);
//			personObj = person;
//			personToCellDataMap.put(person, personToCellDataList);
//
//		}
//		else
//		{
//
//			if(currentName.equalsIgnoreCase(name))
//			{
//				currentRowNum++;
//				for (ExcelCellData excelCellData : cellList) {
//					excelCellData.setRowNum(currentRowNum);
//				}
//				personObj.setEndRow(currentRowNum);
//				personToCellDataMap.get(personObj).addAll(cellList);
//			}
//			else
//			{
//				currentRowNum++;
//				Person person = new Person();
//				person.setName(name);
//				currentName = name;
//				List<ExcelCellData> personToCellDataList = new ArrayList<ExcelCellData>();
//
//
//				personObj.setEndRow(currentRowNum);
//				currentRowNum++;
//
//
//				person.setStartRow(currentRowNum);
//				person.setEndRow(currentRowNum);
//
//				for (ExcelCellData excelCellData : cellList) {
//					excelCellData.setRowNum(currentRowNum);
//
//				}
//
//
//
//				personToCellDataList.addAll(cellList);
//				personObj = person;
//				personToCellDataMap.put(person, personToCellDataList);
//
//
//			}
//		}


	}
	public int getColNum(String colName) {
		for (ExcelCellData cell : headerCells) {
			if(((String)cell.getValue()).equalsIgnoreCase(colName))
			{
				return cell.getColNum();
			}
		}
		return -1;
	}

	public class Person
	{
		private int startRow =-1;
		private int endRow = -1;
		private String Name = null;
		private double totalWorkingHours = 0.0;
		private double diffHours = 0.0;
		List<ExcelCellData> cellDataList = new ArrayList<ExcelCellData>();
		
		public int getStartRow() {
			return startRow;
		}
		public void addCellDataList(List<ExcelCellData> cellList, int currentRowNumber) {

			for (ExcelCellData excelCellData : cellList) {
				excelCellData.setRowNum(currentRowNumber);
			}
			this.cellDataList.addAll(cellList);
			
		}
		public void setStartRow(int startRow) {
			this.startRow = startRow;
		}
		public int getEndRow() {
			return endRow;
		}
		public void setEndRow(int endRow) {
			this.endRow = endRow;
		}
		public String getName() {
			return Name;
		}
		public void setName(String name) {
			Name = name;
		}
	
		@Override
		public boolean equals(Object obj) {
			
			return super.equals(obj);
		}
		@Override
		public int hashCode() {
			
			return super.hashCode();
		}
		public void addCellData(ExcelCellData cell) {
			this.cellDataList.add(cell);
			
		}
		public List<Integer> getRows() {
			List<Integer> rowNumbersList = new ArrayList<Integer>();
			
			for (ExcelCellData cell : cellDataList) {
				int rowNum = cell.getRowNum();
				if(!rowNumbersList.contains(rowNum))
				{
					rowNumbersList.add(rowNum);
				}
				
			}
			return rowNumbersList;
		}
		public List<Integer> getColumns(Integer rowNum) {
			List<Integer> colNumbersList = new ArrayList<Integer>();

			for (ExcelCellData cell : cellDataList) {
				int rowNumber = cell.getRowNum();
				if(rowNumber == rowNum)
				{
					colNumbersList.add(cell.getColNum());
				}

			}
			return colNumbersList;
		}
		public ExcelCellData getExcelCellData(Integer rowNum, Integer colNum) {
			
			for (ExcelCellData cell : cellDataList) {
				int rowNumber = cell.getRowNum();
				int colNumber = cell.getColNum();
				if(rowNumber == rowNum && colNumber == colNum)
				{
					return cell;
				}

			}
			return null;
		}
		public Double getTotalWorkingHours() {
			
			Double totalHours = 0.0;
			for (ExcelCellData cell : cellDataList) {
				if(cell.getColNum() == getColNum(Constants.QUANTITY_COLUMN_NAME))
				{
					if(cell.getRowNum() != endRow)
					{
						Double hours = (Double)cell.getValue();
						totalHours = totalHours +hours;
					}
				}
				
			}
			totalWorkingHours = totalHours;
			return totalHours;
		}
		public List<ExcelCellData> getCellDataList() {
			return cellDataList;
		}
		public void setCellDataList(List<ExcelCellData> cellDataList) {
			this.cellDataList = cellDataList;
		}
		public double getDiffHours() {
			return diffHours;
		}
		public void setDiffHours(double diffHours) {
			this.diffHours = diffHours;
		}
		public void setTotalWorkingHours(double totalWorkingHours) {
			this.totalWorkingHours = totalWorkingHours;
		}
		
	}

	public List<Integer> getStartAndEndRowNums(String name) {
		
		
		List<Integer> startAndEndRowNums =  new ArrayList<Integer>();
		
		for (Person person : personList) {
			
			if(person.getName().equalsIgnoreCase(name))
			{
				int startRow = person.getStartRow();
				int endRow = person.getEndRow();
				startAndEndRowNums.add(startRow);
				startAndEndRowNums.add(endRow);
				break;
			}
		}
		return startAndEndRowNums;
	}

	public int getEndRow(String name) {

		int rowNumber = -1;
		for (Person person : personList) {
			
			if(person.getName().equalsIgnoreCase(name))
			{
				
				 rowNumber = person.getEndRow();
				 break;
				
			}
		}
		return rowNumber;
	}

	public void addCell(String name, ExcelCellData cell) {
		
		
		for (Person person : personList) {

			if(person.getName().equalsIgnoreCase(name))
			{

				person.addCellData(cell);
				break;

			}
		}
	}

	public boolean isNamePresent(String name) {
		for (Person person : personList) {

			if(person.getName().equalsIgnoreCase(name))
			{

				return true;

			}
		}
		return false;
	}

	public Double getTotalWorkingHours(String name) {
		for (Person person : personList) {

			if(person.getName().equalsIgnoreCase(name))
			{
				
				return person.getTotalWorkingHours();
			}
		}
		return -1.0;
	}

	public void setWorkingHoursMatch(boolean b, String name) {
		for (Person person : personList) {

			if(person.getName().equalsIgnoreCase(name))
			{
				
				ExcelCellData excelCellData = person.getExcelCellData(person.getEndRow(), getColNum(Constants.QUANTITY_COLUMN_NAME));
				
				excelCellData.setStyle(CellStyles.NO_FILL_STYLE);
			}
		}
		
	}

	public void setTotalDiffHours(String name,double diffHours) {
		for (Person person : personList) {

			if(person.getName().equalsIgnoreCase(name))
			{
				person.setDiffHours(diffHours);
				break;
				
			}
		}
		
	}

	public int getStartRow(String name) {
		int rowNumber = -1;
		for (Person person : personList) {
			
			if(person.getName().equalsIgnoreCase(name))
			{
				
				 rowNumber = person.getStartRow();
				 break;
				
			}
		}
		return rowNumber;
	}
	
	public Double getDiffHours(String name)
	{
		Double diffHours = 0.0;
		for (Person person : personList) {
			
			if(person.getName().equalsIgnoreCase(name))
			{
				
				 diffHours = person.getDiffHours();
				 break;
				
			}
		}
		return diffHours;
	}

	public void setAdditionalRowForLastName(String fullName) {
		
		for (Person person : personList) {

			if(person.getName().equalsIgnoreCase(fullName))
			{
				int endRow = person.getEndRow();
				endRow++;
				person.setEndRow(endRow);
				break;

			}
		}
	}

	public Person getPerson(String name) {
		
		for (Person person : personList) {
			if(person.getName().equalsIgnoreCase(name))
			{
				return person;
			}
		}
		return null;
	}

	public int getEndColumnNumber() {
		return endColumnNumber;
	}

	public void setEndColumnNumber(int endColumnNumber) {
		this.endColumnNumber = endColumnNumber;
	}
	
}
