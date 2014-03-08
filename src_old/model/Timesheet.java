package model;

import java.util.ArrayList;
import java.util.List;

public class Timesheet {
	
	private String firstName = null;
	private String lasttName = null;
	private String name = null;
	private Double weeklyWorkedHours = -1.0;
	
	private static List<ExcelCellData> headerCells = new ArrayList<ExcelCellData>();
	
	public static List<ExcelCellData> getHeaderCells() {
		return headerCells;
	}
	public void setHeaderCells(List<ExcelCellData> headerCells) {
		Timesheet.headerCells = headerCells;
	}
	
	
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
	public Double getWeeklyWorkedHours() {
		return weeklyWorkedHours;
	}
	public void setWeeklyWorkedHours(Double weeklyWorkedHours) {
		this.weeklyWorkedHours = weeklyWorkedHours;
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
}
