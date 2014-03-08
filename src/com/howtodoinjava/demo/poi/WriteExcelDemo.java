package com.howtodoinjava.demo.poi;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.text.SimpleDateFormat;
import java.util.Iterator;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class WriteExcelDemo 
{
	public static Map<Integer, Object[]> data = new TreeMap<Integer, Object[]>();
	public static Map<String,Double> payRateData = new TreeMap<String, Double>();
	public static Map<String,Double> timeSheetData = new TreeMap<String, Double>();

	private static final int COLUMNSIZE = 8;

	public static void main(String[] args) 
	{
		readFile();
		readPayRate();
		readTimeSheet();
		int totalCount = 1;
		int lastNameColNum = 4;

		int firstNameColNum = 3;
		//Blank workbook
		XSSFWorkbook workbook = new XSSFWorkbook(); 
		
		//Create a blank sheet
		XSSFSheet sheet = workbook.createSheet("Employee Data");
		
		//Iterate over data and write to sheet
		Set<Integer> keyset = data.keySet();
		
		FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
		
		int rownum = 0;
		int rowStart = 2;
		String mainName = null;
		String firstName = null;
		String lastName = null;
		boolean diffInHours = false;
		double differrenceHours = 0;
		for (Integer key : keyset)
		{
		    Row row = sheet.createRow(rownum);
		    Object [] objArr = data.get(key);
		    int cellNum = 0;
		    for (Object obj : objArr)
		    {
		       Cell cell = row.createCell(cellNum);
		       if(obj instanceof String){
		    	   if(rownum != 0)
		    	   {   
		    		   if(cellNum == firstNameColNum )
		    		   {
		    			   firstName = (String)obj;
		    		   }
		    		   if(cellNum == lastNameColNum)
		    		   {
		    			   lastName = (String)obj;
		    			   
		    			   if(mainName == null)
		    			   {
		    				   mainName = firstName +","+lastName;
		    			   }
		    			   else
		    			   {
		    				   String tempName = firstName + "," +((String)obj).trim();
		    				   if(!tempName.equalsIgnoreCase(mainName))
		    				   {
		    					   sheet.shiftRows(rownum, rownum++, 1);
		    					   totalCount++;
		    					   //Cell formulacell = row.createCell(7);
		    					   String addFormula = "SUM"+"("+"H"+rowStart+":"+"H"+ (rownum-1) +")";
		    					   //formulacell.setCellFormula("SUM(H2:H9)");
		    					   rowStart = rownum+1;
		    					   //rownum++;
		    					   //sheet.createRow(rownum);
		    					   //cell = row.createCell(7);
		    					   
		    					   CellStyle style = workbook.createCellStyle();
		    					   style.setFillForegroundColor(HSSFColor.YELLOW.index);
		    					   style.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
		    					   
		    					   XSSFFont font = workbook.createFont();
		    					   font.setBold(true);
		    					   //font.setColor(HSSFColor.RED.index);
		    					   style.setFont(font);
		    					   
		    					   XSSFRow row2 = sheet.createRow(rownum-1);
		    					   XSSFCell cell2 = row2.createCell(7);
		    					   cell2.setCellFormula(addFormula);
		    					   
		    					   int evaluateFormulaCell = evaluator.evaluateFormulaCell(cell2);
		    					   double totalHours = cell2.getNumericCellValue();
		    					   if(totalHours != timeSheetData.get(mainName))
		    					   {
		    						   diffInHours = true;
		    						   cell2.setCellStyle(style);
		    						   if(totalHours > timeSheetData.get(mainName))
		    						   {
		    							   differrenceHours = totalHours - timeSheetData.get(mainName);
		    						   }
		    						   else
		    						   {
		    							   differrenceHours = timeSheetData.get(mainName)- totalHours;
		    						   }
		    					   }
		    					   else
		    					   {
		    						   diffInHours = false;
		    					   }
		    					   
		    					   
		    					   XSSFCell cell3 = row2.createCell(8);
		    					   
		    					   cell3.setCellValue(payRateData.get(mainName));
		    					   if(diffInHours)
		    					   {
		    						   cell3.setCellStyle(style);
		    					   }
		    					   
		    					 
		    					   String mulFormula = "PRODUCT"+"("+"H"+(rownum)+":"+"I"+ (rownum) +")";
		    					   XSSFCell cell4 = row2.createCell(9);
		    					   cell4.setCellFormula(mulFormula);;
		    					   if(diffInHours)
		    					   {
		    						   cell4.setCellStyle(style);
		    					   }
		    					   double tempNumber = cell4.getNumericCellValue();
		    					   
		    					   if(diffInHours)
		    					   {
		    						   XSSFCell cell5 = row2.createCell(10);
		    						   cell5.setCellValue(differrenceHours);
		    						   cell5.setCellStyle(style);
		    						   
		    						   XSSFCell cell6 = row2.createCell(11);
		    						   cell6.setCellValue(timeSheetData.get(mainName));
		    						   cell6.setCellStyle(style);
		    					   }
		    					   
		    					   mainName = firstName +","+lastName;
		    				   }
		    			   }
		    		   }
		    	   }
		           cell.setCellValue((String)obj);
		       }
		       else if(obj instanceof Integer){
		            cell.setCellValue((Integer)obj);
		        }
		       else if(obj instanceof Double){
		            cell.setCellValue((Double)obj);
		        }
		       cellNum++;
		    }
		    rownum++;
		    
		}
		try 
		{
			//Write the workbook in file system
		    FileOutputStream out = new FileOutputStream(new File("Charlie_Haddad_WE_26th_Jan_2014-2_Modify_temp.xlsx"));
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

	private static void readTimeSheet() {
		try
		{
			
			FileInputStream file = new FileInputStream(new File("TimeSheet.xlsx"));

			//Create Workbook instance holding reference to .xlsx file
			XSSFWorkbook workbook = new XSSFWorkbook(file);

			//Get first/desired sheet from the workbook
			XSSFSheet sheet = workbook.getSheetAt(0);

			//Iterate through each rows one by one
			Iterator<Row> rowIterator = sheet.iterator();

			int firstNameCellNum = 0;
			int lastNameCellNum = 1;
			int payRateCellNum = 2;
			
			String firstName = null;
			String lastName = null;
			int rownum = -1;
			while (rowIterator.hasNext()) 
			{
				
				Row row = rowIterator.next();
				rownum++;
				if(rownum >0)
				{
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
							timeSheetData.put(firstName+","+lastName, 0.0);
						}
						Cell payRateCell = row.getCell(payRateCellNum);
						if(payRateCell != null)
						{
							double workingHours = payRateCell.getNumericCellValue();
							Double workingHoursVal = (Double)workingHours;

							if(firstName != null && lastName != null)
							{
								timeSheetData.put(firstName+","+lastName,workingHoursVal);
							}

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

	private static void readPayRate() {
		try
		{
			
			FileInputStream file = new FileInputStream(new File("PayRate.xlsx"));

			//Create Workbook instance holding reference to .xlsx file
			XSSFWorkbook workbook = new XSSFWorkbook(file);

			//Get first/desired sheet from the workbook
			XSSFSheet sheet = workbook.getSheetAt(0);

			//Iterate through each rows one by one
			Iterator<Row> rowIterator = sheet.iterator();

			int firstNameCellNum = 0;
			int lastNameCellNum = 1;
			int payRateCellNum = 2;
			
			String firstName = null;
			String lastName = null;
			int rownum = -1;
			while (rowIterator.hasNext()) 
			{
				
				Row row = rowIterator.next();
				rownum++;
				if(rownum >0)
				{
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
							payRateData.put(firstName+","+lastName, 0.0);
						}
						Cell payRateCell = row.getCell(payRateCellNum);
						if(payRateCell != null)
						{
							double payRate = payRateCell.getNumericCellValue();
							Double dummyPayRate = (Double)payRate;

							if(firstName != null && lastName != null)
							{
								payRateData.put(firstName+","+lastName,dummyPayRate);
							}

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

	private static void readFile() {
		try
		{

			FileInputStream file = new FileInputStream(new File("Charlie_Haddad_WE_26th_Jan_2014-2_Modify.xlsx"));

			//Create Workbook instance holding reference to .xlsx file
			XSSFWorkbook workbook = new XSSFWorkbook(file);

			//Get first/desired sheet from the workbook
			XSSFSheet sheet = workbook.getSheetAt(0);

			//Iterate through each rows one by one
			Iterator<Row> rowIterator = sheet.iterator();
			int rownum = 0;

			while (rowIterator.hasNext()) 
			{
				int cellNum = 0;
				Row row = rowIterator.next();
				//For each row, iterate through all the columns
				Iterator<Cell> cellIterator = row.cellIterator();
				Object[] arr = new Object[COLUMNSIZE];
				while (cellIterator.hasNext()) 
				{
					Cell cell = cellIterator.next();
					
					//Check the cell type and format accordingly
					switch (cell.getCellType()) 
					{
						case Cell.CELL_TYPE_NUMERIC:
							
							if (DateUtil.isCellDateFormatted(cell))
							{
								SimpleDateFormat sdf = new SimpleDateFormat("dd/MM/yyyy HH:mm:ss");
							    String format = sdf.format(cell.getDateCellValue());
							    String[] split = format.split(" ");
							    if(cellNum == 2)
							    {
							    	arr[cellNum] = split[0];
							    }
							    else
							    {
							    	arr[cellNum] = split[1];
							    }
							    
							    
							}
							else
							{
								arr[cellNum] = (Double)cell.getNumericCellValue();
							}
							//System.out.print(cell.getNumericCellValue() + "\t");
							break;
						case Cell.CELL_TYPE_STRING:
							arr[cellNum] = (String)cell.getStringCellValue();
							//System.out.print(cell.getStringCellValue() + "\t");
							break;
						
					}
					cellNum++;
				}
				//System.out.println("");
				data.put(rownum, arr);
				rownum++;
			}
			
			
			
			file.close();
		} 
		catch (Exception e) 
		{
			e.printStackTrace();
		}
		
	}
}
