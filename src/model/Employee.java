package model;

import java.util.ArrayList;
import java.util.List;

public class Employee {

	private String firstName = null;
	private String lasttName = null;
	private String name = null;
	private Double generalRate = -1.0;
	private Double publicHolidayRate = -1.0;
	boolean toBeComapred = false;
	
	
	
	private static List<ExcelCellData> headerCells = new ArrayList<ExcelCellData>();
	
	
	public String getFirstName() {
		return firstName;
	}
	public void setFirstName(String firstName) {
		this.firstName = firstName;
	}
	public String getLasttName() {
		return lasttName;
	}
	public void setLasttName(String lasttName) {
		this.lasttName = lasttName;
	}
	public String getName() {
		return name;
	}
	public void setName(String name) {
		this.name = name;
	}
	public Double getGeneralRate() {
		return generalRate;
	}
	public void setGeneralRate(Double generalRate) {
		this.generalRate = generalRate;
	}
	public Double getPublicHolidayRate() {
		return publicHolidayRate;
	}
	public void setPublicHolidayRate(Double publicHolidayRate) {
		this.publicHolidayRate = publicHolidayRate;
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
	public static List<ExcelCellData> getHeaderCells() {
		return headerCells;
	}
	public void setHeaderCells(List<ExcelCellData> headerCells) {
		this.headerCells = headerCells;
	}
	public boolean isToBeComapred() {
		return toBeComapred;
	}
	public void setToBeComapred(boolean toBeComapred) {
		this.toBeComapred = toBeComapred;
	}
	
}
