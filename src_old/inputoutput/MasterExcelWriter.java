package inputoutput;

import java.io.File;
import java.io.FileOutputStream;
import java.util.List;

import model.CellStyles;
import model.Constants;
import model.Employee;
import model.ExcelCellData;
import model.ExcelSheet;
import model.ExcelSheet.Person;
import model.Invoice;
import model.Model;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.format.CellNumberFormatter;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class MasterExcelWriter implements IExcelFileWriter{

	@Override
	public void write(Model model, String destFile) {
		
		XSSFWorkbook workbook = new XSSFWorkbook(); 
		 
		
		
		
		//Create a blank sheet
		XSSFSheet sheet = workbook.createSheet("Employee Data");
		
		ExcelSheet masterSheet = model.getSheet();
		
		
		List<Person> personList = masterSheet.getPersonList();
		for (Person person : personList) {
			List<Integer> rowList = person.getRows();
			for (Integer rowNum : rowList) {
				Row row = sheet.createRow(rowNum);
				List<Integer> columnList = person.getColumns(rowNum);
				for (Integer colNum : columnList) {
					Cell cell = row.createCell(colNum);
					ExcelCellData excelCellData = person.getExcelCellData(rowNum,colNum);
					Object obj = excelCellData.getValue();

					if(excelCellData.isSpecialCell())
					{
						cell.setCellFormula(excelCellData.getFormula());
						
					}
					else{
						if(obj instanceof String)
						{
							cell.setCellValue((String)obj);

						}
						else if(obj instanceof Integer){
							cell.setCellValue((Integer)obj);
						}
						else if(obj instanceof Double){
							cell.setCellValue((Double)obj);
						}
					 }
					if(excelCellData.isSetStyle())
					{

						if(excelCellData.getStyle() == CellStyles.NO_FILL_STYLE)
						{
							
						}
						else{
							CellStyle style = workbook.createCellStyle();
							style.setFillForegroundColor(HSSFColor.YELLOW.index);
							style.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);

							XSSFFont font = workbook.createFont();
							font.setBold(true);
							//font.setColor(HSSFColor.RED.index);
							style.setFont(font);
							cell.setCellStyle(style);
						}
					}
				}
			}
		}
		
		
//		Map<Person, List<ExcelCellData>> personToCellDataMap = masterSheet.getPersonToCellDataMap();
//		Set<Person> keySet = personToCellDataMap.keySet();
//		for (Person person : keySet) {
//			List<ExcelCellData> list = personToCellDataMap.get(person);
//			int rowNum = -1;
//			for (ExcelCellData excelCellData : list) {
//				Row createRow = null;
//				Row row = null;
//				if(rowNum == -1)
//				{
//					rowNum = excelCellData.getRowNum();
//					row = sheet.createRow(excelCellData.getRowNum());
//				}
//				if(excelCellData.getRowNum() == rowNum)
//				{
//					
//					
//				}
//				else
//				{
//					row = sheet.createRow(excelCellData.getRowNum());
//					rowNum = excelCellData.getRowNum();
//				}
//				 
//				 Cell cell = row.createCell(excelCellData.getColNum());
// 				 Object obj = excelCellData.getValue();
//				 
//				 if(excelCellData.isSpecialCell())
//				 {
//					 cell.setCellFormula(excelCellData.getFormula());
//					 if(excelCellData.isSetStyle())
//					 {
//						 
//						 CellStyle style = workbook.createCellStyle();
//						 style.setFillForegroundColor(HSSFColor.YELLOW.index);
//						 style.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
//
//						 XSSFFont font = workbook.createFont();
//						 font.setBold(true);
//						 //font.setColor(HSSFColor.RED.index);
//						 style.setFont(font);
//						 cell.setCellStyle(style);
//					 }
//				 }
//				 else{
//					 if(obj instanceof String)
//					 {
//						 cell.setCellValue((String)obj);
//
//					 }
//					 else if(obj instanceof Integer){
//						 cell.setCellValue((Integer)obj);
//					 }
//					 else if(obj instanceof Double){
//						 cell.setCellValue((Double)obj);
//					 }
//				 }
//
//			}
//		}
		
		writeNegativeDiffSheet(model,workbook);
		writePositiveDiffSheet(model,workbook);
		writeMissingNames(model,workbook);
		writeNotComparedNames(model,workbook);
		//writeInterTimeSheet(model,workbook);
		writeNotComparedInvoices(model,workbook);
		writeGeneralInvoices(model,workbook);
		try 
		{
			//Write the workbook in file system
		    FileOutputStream out = new FileOutputStream(new File(destFile));
		    workbook.write(out);
		    out.close();
		  
		    //System.out.println("howtodoinjava_demo.xlsx written successfully on disk.");
		     
		} 
		catch (Exception e) 
		{
		    e.printStackTrace();
		}
	}

	private void writeGeneralInvoices(Model model, XSSFWorkbook workbook) {
		List<Invoice> invoiceList = model.getInvoiceList();
		int rowNum = 0;
		int colNum = 0;
		
		XSSFSheet sheet = workbook.createSheet("Invoice_Notadjusted");
		
		for (Invoice invoice : invoiceList) {
			String name = invoice.getFullName();
			Row row = sheet.createRow(rowNum);
			rowNum++;
			String[] split = name.split(",");
			if((split != null) && (split.length >= 2))
			{
				
				for (String string : split) {
					
					Cell cell = row.createCell(colNum);
					colNum++;
					cell.setCellValue(string);
			
				}
			}
			Cell cell = row.createCell(colNum);
			cell.setCellValue(invoice.getFinalRate());
			
			colNum = 0;
			
		}
		
	}

	private void writeNotComparedInvoices(Model model, XSSFWorkbook workbook) {
	
		List<Invoice> notComparedInvoiceList = model.getNotComparedInvoiceList();
		int rowNum = 0;
		int colNum = 0;
		
		XSSFSheet sheet = workbook.createSheet("Invoice_NotComparedNamesList");
		
		for (Invoice invoice : notComparedInvoiceList) {
			String name = invoice.getFullName();
			Row row = sheet.createRow(rowNum);
			rowNum++;
			String[] split = name.split(",");
			if((split != null) && (split.length >= 2))
			{
				
				for (String string : split) {
					
					Cell cell = row.createCell(colNum);
					colNum++;
					cell.setCellValue(string);
			
				}
			}
			Cell cell = row.createCell(colNum);
			cell.setCellValue(invoice.getFinalRate());
			
			colNum = 0;
			
		}
		
	}

	private void writeInterTimeSheet(Model model, XSSFWorkbook workbook) {
		
		List<String> missingNames = model.getMissingNames();
		XSSFSheet sheet = workbook.createSheet("People Who Worked This Week");
		int rowNum = 0;
		int colNum = 0;
		
		
		ExcelSheet workingSheet = model.getSheet();
		List<Person> personList = workingSheet.getPersonList();
		for (Person person : personList) {
			String name = person.getName();
			if(model.isEmployeeToBeCompared(name))
			{

				Row row = sheet.createRow(rowNum);
				rowNum++;
				String[] split = name.split(",");
				if((split != null) && (split.length >= 2))
				{

					for (String string : split) {

						Cell cell = row.createCell(colNum);
						colNum++;
						cell.setCellValue(string);

					}
				}
				colNum = 0;
			}
		}
		
	}

	private void writeMissingNames(Model model, XSSFWorkbook workbook) {
		
		List<String> missingNames = model.getMissingNames();
		int rowNum = 0;
		int colNum = 0;
		
		XSSFSheet sheet = workbook.createSheet("Missing Hours Names List");
		
		for (String name : missingNames) {
			Row row = sheet.createRow(rowNum);
			rowNum++;
			String[] split = name.split(",");
			if((split != null) && (split.length >= 2))
			{
			
				for (String string : split) {
					
					Cell cell = row.createCell(colNum);
					colNum++;
					cell.setCellValue(string);
					
				}
			}
			colNum = 0;
			
		}
		
//		for (String name : missingNames) {
//			Row row = sheet.createRow(rowNum);
//			rowNum++;
//			Cell cell = row.createCell(colNum);
//			//colNum++;
//			cell.setCellValue(name);
//		}
	}
	
private void writeNotComparedNames(Model model, XSSFWorkbook workbook) {
		
		List<String> peopleNotComparedNames = model.getPeopleNotCompared();
		int rowNum = 0;
		int colNum = 0;
		
		XSSFSheet sheet = workbook.createSheet("Not Compared Names List");
		
		for (String name : peopleNotComparedNames) {
			Row row = sheet.createRow(rowNum);
			rowNum++;
			String[] split = name.split(",");
			if((split != null) && (split.length >= 2))
			{
				
				for (String string : split) {
					
					Cell cell = row.createCell(colNum);
					colNum++;
					cell.setCellValue(string);
			
				}
			}
			colNum = 0;
			
		}
	}

	private void writeNegativeDiffSheet(Model model, XSSFWorkbook workbook) {
		
		int rowNum = 0;
		int colNum = 0;
		//int endColumnNumber = model.getSheet().getColNum(Constants.QUANTITY_COLUMN_NAME);
		int endColumnNumber = model.getSheet().getEndColumnNumber();
		
		List<String> diffInWorkingHoursNames = model.getDiffInWorkingHoursNames();
		XSSFSheet sheet = workbook.createSheet("Negative Hours Data");
		ExcelSheet masterSheet = model.getSheet();

		List<Person> personList = masterSheet.getPersonList();

		for (String  personName : diffInWorkingHoursNames) {

			Double diffHours = masterSheet.getDiffHours(personName);
			if(diffHours < 0){
				List<ExcelCellData> cellDataList =getCellList(personList,personName);

				if(cellDataList != null)
				{
					colNum = 0;
					Row row = sheet.createRow(rowNum);
					rowNum++;

					for(ExcelCellData excelCellData : cellDataList)
					{
						Cell cell = row.createCell(colNum);
						colNum++;
						Object obj = excelCellData.getValue();

						if(obj == null && excelCellData.isSpecialCell())
						{
							String formula = excelCellData.getFormula();
							//cell.setCellFormula(formula);
						}
						else
						{
							if(obj instanceof String)
							{
								cell.setCellValue((String)obj);

							}
							else if(obj instanceof Integer){
								cell.setCellValue((Integer)obj);
							}
							else if(obj instanceof Double){
								cell.setCellValue((Double)obj);
							}
							if(colNum == (endColumnNumber + 1))
							{

								row = sheet.createRow(rowNum);
								rowNum++;
								colNum = 0;
							}
						}


					}



					Cell cell = row.createCell(endColumnNumber+1);
					cell.setCellValue((Double)diffHours);

					CellStyle style = workbook.createCellStyle();
					style.setFillForegroundColor(HSSFColor.YELLOW.index);
					style.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);

					XSSFFont font = workbook.createFont();
					font.setBold(true);
					//font.setColor(HSSFColor.RED.index);
					style.setFont(font);
					cell.setCellStyle(style);


				}

			}
		}

	}

	private void writePositiveDiffSheet(Model model, XSSFWorkbook workbook) {

		int rowNum = 0;
		int colNum = 0;
		//int endColumnNumber = model.getSheet().getColNum(Constants.QUANTITY_COLUMN_NAME);
		int endColumnNumber = model.getSheet().getEndColumnNumber();
		
		List<String> diffInWorkingHoursNames = model.getDiffInWorkingHoursNames();
		XSSFSheet sheet = workbook.createSheet("Positive Hours Data");
		ExcelSheet masterSheet = model.getSheet();
		
		List<Person> personList = masterSheet.getPersonList();
		
		for (String  personName : diffInWorkingHoursNames) {


			Double diffHours = masterSheet.getDiffHours(personName);

			if(diffHours > 0)
			{


				List<ExcelCellData> cellDataList =getCellList(personList,personName);
				if(cellDataList != null)
				{
					colNum = 0;
					Row row = sheet.createRow(rowNum);
					rowNum++;

					for(ExcelCellData excelCellData : cellDataList)
					{
						Cell cell = row.createCell(colNum);
						colNum++;
						Object obj = excelCellData.getValue();

						if(obj == null && excelCellData.isSpecialCell())
						{
							String formula = excelCellData.getFormula();
							//cell.setCellFormula(formula);
						}
						else
						{
							if(obj instanceof String)
							{
								cell.setCellValue((String)obj);

							}
							else if(obj instanceof Integer){
								cell.setCellValue((Integer)obj);
							}
							else if(obj instanceof Double){
								cell.setCellValue((Double)obj);
							}
							if(colNum == (endColumnNumber + 1))
							{

								row = sheet.createRow(rowNum);
								rowNum++;
								colNum = 0;
							}
						}


					}

					Cell cell = row.createCell(endColumnNumber+1);
					cell.setCellValue((Double)diffHours);

					CellStyle style = workbook.createCellStyle();
					style.setFillForegroundColor(HSSFColor.YELLOW.index);
					style.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);

					XSSFFont font = workbook.createFont();
					font.setBold(true);
					//font.setColor(HSSFColor.RED.index);
					style.setFont(font);
					cell.setCellStyle(style);


				}

			}
		}

	}
	private List<ExcelCellData> getCellList(List<Person> personList,
			String personName) {
		for (Person person : personList) {
			if(person.getName().equalsIgnoreCase(personName))
			{
				return person.getCellDataList();
			}
		}
		return null;
	}

}
