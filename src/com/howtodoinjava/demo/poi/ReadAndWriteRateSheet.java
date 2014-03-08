package com.howtodoinjava.demo.poi;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.Iterator;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class ReadAndWriteRateSheet 
{
	public static Map<String,Double> data = new TreeMap<String, Double>();


	public static void main(String[] args) 
	{
		readFile();
		int totalCount = 1;
		
		//Blank workbook
		XSSFWorkbook workbook = new XSSFWorkbook(); 
		
		//Create a blank sheet
		XSSFSheet sheet = workbook.createSheet("Pay Rate");
		 
		//Iterate over data and write to sheet
		Set<String> keySet = data.keySet();
		
		int rowNum = 0;
		
		
		Row row = sheet.createRow(rowNum);
		createHeader(row);
		rowNum++;
		
		String firstName = null;
		String lastName = null;
		
		for (String key : keySet)
		{
			int cellNum = 0;
			row = sheet.createRow(rowNum);
		    Double payRate = data.get(key);
		    
		    String[] names = key.split(",");
		   
		    if(names != null && names[0] != null && names[1] != null)
		    {
		    	firstName = names[0];
		    	lastName = names[1]; 
		    }
	   
		    Cell cell = row.createCell(cellNum);
		    
		    cell = row.createCell(cellNum);
			cell.setCellValue(firstName);
			cellNum++;
			
			cell = row.createCell(cellNum);
			cell.setCellValue(lastName);
			
			cellNum++;
			cell = row.createCell(cellNum);
			cell.setCellValue(payRate);
			


		    rowNum++;
		    
		}
		try 
		{
			//Write the workbook in file system
		    FileOutputStream out = new FileOutputStream(new File("PayRate.xlsx"));
		    workbook.write(out);
		    out.close();
		    System.out.println(totalCount);
		    //System.out.println("howtodoinjava_demo.xlsx written successfully on disk.");
		     
		} 
		catch (Exception e) 
		{
		    e.printStackTrace();
		}
	}

	private static void createHeader(Row row) {
		int cellNum = 0;
		Cell cell = row.createCell(cellNum);
		cell.setCellValue("First Name");
		cellNum++;
		cell = row.createCell(cellNum);
		cell.setCellValue("Last Name");
		
		cellNum++;
		cell = row.createCell(cellNum);
		cell.setCellValue("Pay Rate");
		
	}

	private static void readFile() {
		try
		{
			
			FileInputStream file = new FileInputStream(new File("pav.xlsx"));

			//Create Workbook instance holding reference to .xlsx file
			XSSFWorkbook workbook = new XSSFWorkbook(file);

			//Get first/desired sheet from the workbook
			XSSFSheet sheet = workbook.getSheetAt(0);

			//Iterate through each rows one by one
			Iterator<Row> rowIterator = sheet.iterator();

			int firstNameCellNum = 3;
			int lastNameCellNum = 4;
			int payRateCellNum = 8;
			
			String firstName = null;
			String lastName = null;
			
			while (rowIterator.hasNext()) 
			{
				
				Row row = rowIterator.next();
				
				//For each row, iterate through all the columns
				if(row.getCell(0)!= null)
				{
					Cell firstNameCell = row.getCell(firstNameCellNum);
					if(firstNameCell != null)
					{
						firstName = firstNameCell.getStringCellValue();

					}
					Cell lastNameCell = row.getCell(lastNameCellNum);
					if(lastNameCell != null)
					{
						lastName = lastNameCell.getStringCellValue();

					}
					if(firstName != null && lastName != null)
					{
						data.put(firstName+","+lastName, 0.0);
					}
				}
				else
				{
					Cell payRateCell = row.getCell(payRateCellNum);
					if(payRateCell != null)
					{
						double payRate = payRateCell.getNumericCellValue();
						Double dummyPayRate = (Double)payRate;
						
						if(firstName != null && lastName != null)
						{
							data.put(firstName+","+lastName,dummyPayRate);
						}
						
					}
				}
												
			
			}
			
			
			
			file.close();
		} 
		catch (Exception e) 
		{
			e.printStackTrace();
		}
		
	}
}
