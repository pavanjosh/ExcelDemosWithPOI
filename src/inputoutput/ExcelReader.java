package inputoutput;

import java.io.File;
import java.io.FileInputStream;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.TreeMap;

import model.Constants;
import model.Employee;
import model.ExcelCellData;
import model.ExcelSheet;
import model.Model;
import model.Timesheet;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelReader implements IReader{

	
	
	
	public static Map<Integer, Object[]> data = new TreeMap<Integer, Object[]>();
	int dateCellNum = 2;
	@Override
	public Model read(Map<String,String> fileList) {
		
		Model model = null;
		try
		{
			String masterSheetFilePath = fileList.get("MasterDataSheet");
			String payRateSheetPath = fileList.get("PayRateDataSheet");
			String timeSheetPath = fileList.get("TimeSheetDataSheet");
			String missedShiftPath = fileList.get("MissedShiftData");
			
			List<Employee> employeeList = null;
			List<Timesheet> timeSheetList = null;
			ExcelSheet masterSheet = null;
			ExcelSheet missedShiftSheet = null;
			
			
			
			if(masterSheetFilePath != null)
			{
				masterSheet = readMasterDataSheet(masterSheetFilePath);
			}
			if(payRateSheetPath != null)
			{
				employeeList = readPayRateDataSheet(payRateSheetPath);
			}
			
			if(timeSheetPath != null)
			{
				timeSheetList =  readTimeSheetDataSheet(timeSheetPath);
			}
			
			if(missedShiftPath != null)
			{
				missedShiftSheet = readMissedShiftData(missedShiftPath);
			}
			
			model = new Model();
			model.setEmployeeList(employeeList);
			model.setTimeSheetList(timeSheetList);
			model.setSheet(masterSheet);
			model.setMissedShiftsheet(missedShiftSheet);
			

		} 
		catch (Exception e) 
		{
			e.printStackTrace();
		}
	
		
		return model;
		
	}

	private ExcelSheet readMissedShiftData(String missedShiftPath) {
		ExcelSheet missedShiftSheet = null;
		try
		{

			String firstName = null;
			String lastName = null;
			String fullName = null;
			
			FileInputStream file = new FileInputStream(new File(missedShiftPath));

			//Create Workbook instance holding reference to .xlsx file
			XSSFWorkbook workbook = new XSSFWorkbook(file);

			//Get first/desired sheet from the workbook
			XSSFSheet sheet = workbook.getSheetAt(0);
			missedShiftSheet = new ExcelSheet();
			
			List<ExcelCellData> headerCells = missedShiftSheet.getHeaderCells();
			
			//Iterate through each rows one by one
			Iterator<Row> rowIterator = sheet.iterator();
			int excelSheetRowNum = 0;
			int dataStructureRowNum = 0;
			List<ExcelCellData> tempCellDataList = new ArrayList<ExcelCellData>();
			while (rowIterator.hasNext()) 
			{
				tempCellDataList.clear();
				int cellNum = 0;
				Row row = rowIterator.next();
				//For each row, iterate through all the columns
				Iterator<Cell> cellIterator = row.cellIterator();
				//Object[] arr = new Object[COLUMNSIZE];
				while (cellIterator.hasNext()) 
				{
					Cell cell = cellIterator.next();
					ExcelCellData cellData = new ExcelCellData();
					cellData.setColNum(cellNum);
					cellData.setRowNum(dataStructureRowNum);
					
					// Reading all header cells
					// Assumed header cells are in row 0
					// Also it is assumed that all the header cells have string values
					if(excelSheetRowNum == 0)
					{
						ExcelCellData headerCellData = new ExcelCellData();
						headerCellData.setColNum(cellNum);
						headerCellData.setRowNum(dataStructureRowNum);
						String headerName = (String)cell.getStringCellValue();
						headerCellData.setValue(headerName);
						headerCells.add(headerCellData);

					}
					else
					{
						//Check the cell type and format accordingly
						switch (cell.getCellType()) 
						{
						case Cell.CELL_TYPE_NUMERIC:


							if (DateUtil.isCellDateFormatted(cell))
							{

								SimpleDateFormat sdf = new SimpleDateFormat("dd/MM/yyyy HH:mm:ss");
								String format = sdf.format(cell.getDateCellValue());
								String[] split = format.split(" ");
								if(cellNum == missedShiftSheet.getColNum(Constants.DATE_COLUMN_NAME))
								{
									//arr[cellNum] = split[0];
									cellData.setValue(split[0]);
								}
								else
								{
									//arr[cellNum] = split[1];
									cellData.setValue(split[1]);
								}


							}
							else
							{
								//arr[cellNum] = (Double)cell.getNumericCellValue();
								
								cellData.setValue((Double)cell.getNumericCellValue());
								
							}
							//System.out.print(cell.getNumericCellValue() + "\t");
							break;
						case Cell.CELL_TYPE_STRING:
							//arr[cellNum] = (String)cell.getStringCellValue();
							if(cellNum == missedShiftSheet.getColNum(Constants.FIRSTNAME_COLUMN_NAME))
							{
								firstName = (String)cell.getStringCellValue();
							}
							else if(cellNum == missedShiftSheet.getColNum(Constants.LASTNAME_COLUMN_NAME))
							{
								lastName = (String)cell.getStringCellValue();
							}
							if(firstName != null && lastName != null)
							{
								fullName = firstName+","+lastName;
								firstName = null;
								lastName = null;
							}
							cellData.setValue((String)cell.getStringCellValue());
							//System.out.print(cell.getStringCellValue() + "\t");
							break;

						}
						tempCellDataList.add(cellData);
					}
					
					cellNum++;
					
					
				}
				//System.out.println("");
				//data.put(rowNum, arr);
				if(excelSheetRowNum != 0)
				{
					missedShiftSheet.populateSheetData(fullName, tempCellDataList);
				}
				
				excelSheetRowNum++;
				
				dataStructureRowNum++;
			}
			missedShiftSheet.setAdditionalRowForLastName(fullName);
			file.close();
		} 
		catch (Exception e) 
		{
			missedShiftSheet = null;
			e.printStackTrace();
		}

		return missedShiftSheet;
		
	}

	private List<Timesheet> readTimeSheetDataSheet(String timeSheetPath) {

		List<Timesheet> timeSheetList = null;
		try
		{

			String firstName = null;
			String lastName = null;
			String fullName = null;
			
			FileInputStream file = new FileInputStream(new File(timeSheetPath));

			//Create Workbook instance holding reference to .xlsx file
			XSSFWorkbook workbook = new XSSFWorkbook(file);

			//Get first/desired sheet from the workbook
			XSSFSheet sheet = workbook.getSheetAt(0);
			
			
			timeSheetList = new ArrayList<Timesheet>();
			
			
			//Iterate through each rows one by one
			Iterator<Row> rowIterator = sheet.iterator();
			int excelSheetRowNum = 0;
			int dataStructureRowNum = 0;
			List<ExcelCellData> tempCellDataList = new ArrayList<ExcelCellData>();
			while (rowIterator.hasNext()) 
			{
				tempCellDataList.clear();
				int cellNum = 0;
				Row row = rowIterator.next();
				Timesheet timeSheet = new Timesheet();
				//For each row, iterate through all the columns
				Iterator<Cell> cellIterator = row.cellIterator();
				//Object[] arr = new Object[COLUMNSIZE];
				while (cellIterator.hasNext()) 
				{
					Cell cell = cellIterator.next();
					
					
					List<ExcelCellData> headerCells = Timesheet.getHeaderCells();
					
					// Reading all header cells
					// Assumed header cells are in row 0
					// Also it is assumed that all the header cells have string values
					if(excelSheetRowNum == 0)
					{
						ExcelCellData headerCellData = new ExcelCellData();
						headerCellData.setColNum(cellNum);
						headerCellData.setRowNum(dataStructureRowNum);
						String headerName = (String)cell.getStringCellValue();
						headerCellData.setValue(headerName);
						headerCells.add(headerCellData);
					}
					else
					{
						//Check the cell type and format accordingly
						switch (cell.getCellType()) 
						{
						case Cell.CELL_TYPE_NUMERIC:
							if(cellNum == timeSheet.getColNum(Constants.TOTAL_WORKED_HOURS_COLUMN_NAME))
							{
								timeSheet.setWeeklyWorkedHours((Double)cell.getNumericCellValue());
							}
							
							
							break;
						case Cell.CELL_TYPE_STRING:
							//arr[cellNum] = (String)cell.getStringCellValue();
							if(cellNum == timeSheet.getColNum(Constants.FIRSTNAME_COLUMN_NAME))
							{
								firstName = (String)cell.getStringCellValue();
								timeSheet.setFirstName(firstName);
							}
							else if(cellNum == timeSheet.getColNum(Constants.LASTNAME_COLUMN_NAME))
							{
								lastName = (String)cell.getStringCellValue();
								timeSheet.setLasttName(lastName);
							}
							else if(cellNum == timeSheet.getColNum(Constants.TOTAL_WORKED_HOURS_COLUMN_NAME))
							{
								timeSheet.setWeeklyWorkedHours(-1000.00);
							}
							if(firstName != null && lastName != null)
							{
								fullName = firstName+","+lastName;
								firstName = null;
								lastName = null;
							}
							if(fullName != null)
							{
								timeSheet.setName(fullName);
								fullName= null;
							}
							//System.out.print(cell.getStringCellValue() + "\t");
							break;

						}
						
					}
					
					cellNum++;
					
					
				}
				//System.out.println("");
				//data.put(rowNum, arr);
				if(excelSheetRowNum != 0)
				{
					timeSheetList.add(timeSheet);
				}
				excelSheetRowNum++;
				
				dataStructureRowNum++;
			}

			file.close();
		} 
		catch (Exception e) 
		{
			timeSheetList = null;
			e.printStackTrace();
		}
		return timeSheetList;
		
	}

	private List<Employee> readPayRateDataSheet(String payRateSheetPath) {

		List<Employee> employeeList = null;
		try
		{

			String firstName = null;
			String lastName = null;
			String fullName = null;
			String toBeComapred = null;
			
			FileInputStream file = new FileInputStream(new File(payRateSheetPath));

			//Create Workbook instance holding reference to .xlsx file
			XSSFWorkbook workbook = new XSSFWorkbook(file);

			//Get first/desired sheet from the workbook
			XSSFSheet sheet = workbook.getSheetAt(0);
			
			employeeList = new ArrayList<Employee>();
			
			
			
			//Iterate through each rows one by one
			Iterator<Row> rowIterator = sheet.iterator();
			int excelSheetRowNum = 0;
			int dataStructureRowNum = 0;
			List<ExcelCellData> tempCellDataList = new ArrayList<ExcelCellData>();
			while (rowIterator.hasNext()) 
			{
				tempCellDataList.clear();
				int cellNum = 0;
				Row row = rowIterator.next();
				Employee empl = new Employee();
				//For each row, iterate through all the columns
				Iterator<Cell> cellIterator = row.cellIterator();
				//Object[] arr = new Object[COLUMNSIZE];
				while (cellIterator.hasNext()) 
				{
					Cell cell = cellIterator.next();
					
					
					List<ExcelCellData> headerCells = Employee.getHeaderCells();
					
					// Reading all header cells
					// Assumed header cells are in row 0
					// Also it is assumed that all the header cells have string values
					if(excelSheetRowNum == 0)
					{
						ExcelCellData headerCellData = new ExcelCellData();
						headerCellData.setColNum(cellNum);
						headerCellData.setRowNum(dataStructureRowNum);
						String headerName = (String)cell.getStringCellValue();
						headerCellData.setValue(headerName);
						headerCells.add(headerCellData);
					}
					else
					{
						//Check the cell type and format accordingly
						switch (cell.getCellType()) 
						{
						case Cell.CELL_TYPE_NUMERIC:
							if(cellNum == empl.getColNum(Constants.PAYRATE_COLUMN_NAME))
							{
								empl.setGeneralRate((Double)cell.getNumericCellValue());
							}
							else if(cellNum == empl.getColNum(Constants.PUBLICHOLIDAY_PAYRATE_COLUMN_NAME)){
								empl.setPublicHolidayRate((Double)cell.getNumericCellValue());
							}
							
							break;
						case Cell.CELL_TYPE_STRING:
							//arr[cellNum] = (String)cell.getStringCellValue();
							if(cellNum == empl.getColNum(Constants.FIRSTNAME_COLUMN_NAME))
							{
								firstName = (String)cell.getStringCellValue();
								empl.setFirstName(firstName);
							}
							else if(cellNum == empl.getColNum(Constants.LASTNAME_COLUMN_NAME))
							{
								lastName = (String)cell.getStringCellValue();
								empl.setLasttName(lastName);
							}
							else if(cellNum == empl.getColNum(Constants.TO_BE_COMPARED_NAME))
							{
								toBeComapred = (String)cell.getStringCellValue();
								if(toBeComapred.equalsIgnoreCase("n"))
								{
									empl.setToBeComapred(false);
								}
								else
								{
									empl.setToBeComapred(true);
								}
									
							}
							if(firstName != null && lastName != null)
							{
								fullName = firstName+","+lastName;
								firstName = null;
								lastName = null;
							}
							if(fullName != null)
							{
								empl.setName(fullName);
								fullName= null;
							}
							//System.out.print(cell.getStringCellValue() + "\t");
							break;

						}
						
					}
					
					cellNum++;
					
					
				}
				//System.out.println("");
				//data.put(rowNum, arr);
				if(excelSheetRowNum != 0)
				{
					employeeList.add(empl);
				}
				excelSheetRowNum++;
				
				dataStructureRowNum++;
			}

			file.close();
		} 
		catch (Exception e) 
		{
			employeeList = null;
			e.printStackTrace();
		}
		
		return employeeList;
	}

	private ExcelSheet readMasterDataSheet(String masterSheetFilePath) {
		ExcelSheet masterSheet = null;
		try
		{

			String firstName = null;
			String lastName = null;
			String fullName = null;
			
			FileInputStream file = new FileInputStream(new File(masterSheetFilePath));

			//Create Workbook instance holding reference to .xlsx file
			XSSFWorkbook workbook = new XSSFWorkbook(file);

			//Get first/desired sheet from the workbook
			XSSFSheet sheet = workbook.getSheetAt(0);
			masterSheet = new ExcelSheet();
			
			List<ExcelCellData> headerCells = masterSheet.getHeaderCells();
			
			//Iterate through each rows one by one
			Iterator<Row> rowIterator = sheet.iterator();
			int excelSheetRowNum = 0;
			int dataStructureRowNum = 0;
			List<ExcelCellData> tempCellDataList = new ArrayList<ExcelCellData>();
			while (rowIterator.hasNext()) 
			{
				tempCellDataList.clear();
				int cellNum = 0;
				Row row = rowIterator.next();
				//For each row, iterate through all the columns
				Iterator<Cell> cellIterator = row.cellIterator();
				//Object[] arr = new Object[COLUMNSIZE];
				while (cellIterator.hasNext()) 
				{
					Cell cell = cellIterator.next();
					ExcelCellData cellData = new ExcelCellData();
					cellData.setColNum(cellNum);
					cellData.setRowNum(dataStructureRowNum);
					
					// Reading all header cells
					// Assumed header cells are in row 0
					// Also it is assumed that all the header cells have string values
					if(excelSheetRowNum == 0)
					{
						ExcelCellData headerCellData = new ExcelCellData();
						headerCellData.setColNum(cellNum);
						headerCellData.setRowNum(dataStructureRowNum);
						String headerName = (String)cell.getStringCellValue();
						headerCellData.setValue(headerName);
						headerCells.add(headerCellData);

					}
					else
					{
						//Check the cell type and format accordingly
						switch (cell.getCellType()) 
						{
						case Cell.CELL_TYPE_NUMERIC:


							if (DateUtil.isCellDateFormatted(cell))
							{

								SimpleDateFormat sdf = new SimpleDateFormat("dd/MM/yyyy HH:mm:ss");
								String format = sdf.format(cell.getDateCellValue());
								String[] split = format.split(" ");
								if(cellNum == masterSheet.getColNum(Constants.DATE_COLUMN_NAME))
								{
									//arr[cellNum] = split[0];
									cellData.setValue(split[0]);
								}
								else
								{
									//arr[cellNum] = split[1];
									cellData.setValue(split[1]);
								}


							}
							else
							{
								//arr[cellNum] = (Double)cell.getNumericCellValue();
								
								cellData.setValue((Double)cell.getNumericCellValue());
								
							}
							//System.out.print(cell.getNumericCellValue() + "\t");
							break;
						case Cell.CELL_TYPE_STRING:
							//arr[cellNum] = (String)cell.getStringCellValue();
							if(cellNum == masterSheet.getColNum(Constants.FIRSTNAME_COLUMN_NAME))
							{
								firstName = (String)cell.getStringCellValue();
							}
							else if(cellNum == masterSheet.getColNum(Constants.LASTNAME_COLUMN_NAME))
							{
								lastName = (String)cell.getStringCellValue();
							}
							if(firstName != null && lastName != null)
							{
								fullName = firstName+","+lastName;
								firstName = null;
								lastName = null;
							}
							cellData.setValue((String)cell.getStringCellValue());
							//System.out.print(cell.getStringCellValue() + "\t");
							break;

						}
						tempCellDataList.add(cellData);
					}
					
					cellNum++;
					
					
				}
				//System.out.println("");
				//data.put(rowNum, arr);
				if(excelSheetRowNum != 0)
				{
					masterSheet.populateSheetData(fullName, tempCellDataList);
				}
				else
				{
					cellNum--;
					masterSheet.setEndColumnNumber(cellNum);
				}
				
				excelSheetRowNum++;
				
				dataStructureRowNum++;
			}
			masterSheet.setAdditionalRowForLastName(fullName);
			file.close();
		} 
		catch (Exception e) 
		{
			masterSheet = null;
			e.printStackTrace();
		}

		return masterSheet;
	}

}
