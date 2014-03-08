package businesslogic;

import inputoutput.ExcelReader;
import inputoutput.IReader;
import inputoutput.IWriter;
import inputoutput.MasterExcelWriter;

import java.util.Map;
import java.util.TreeMap;

import model.Model;

public class ProgramManager {

	public static void main(String args[])
	{
		System.out.println("Hello");
		Map<String,String> sourceFileList = new TreeMap<String,String>();
		
		
		//args[0] = "Charlie_Haddad_WE_26th_Jan_2014-2_Modify";
//		sourceFileList.put("MasterDataSheet","E:\\Charlie_Haddad\\demos\\ExcelDemosWithPOI_redesigned\\ExcelDemosWithPOI\\Charlie_Haddad_WE_26th_Jan_2014-2_Modify.xlsx");
//		sourceFileList.put("PayRateDataSheet","E:\\Charlie_Haddad\\demos\\ExcelDemosWithPOI_redesigned\\ExcelDemosWithPOI\\PayRate.xlsx");
//		sourceFileList.put("TimeSheetDataSheet","E:\\Charlie_Haddad\\demos\\ExcelDemosWithPOI_redesigned\\ExcelDemosWithPOI\\TimeSheet.xlsx");
//		
//		String destFileName = "E:\\Charlie_Haddad\\demos\\ExcelDemosWithPOI_redesigned\\ExcelDemosWithPOI\\Charlie_Haddad_WE_26th_Jan_2014-2_Modify_pavan.xlsx";
//		
		
		sourceFileList.put("MasterDataSheet","E:\\Charlie_Haddad\\demos\\ExcelDemosWithPOI_redesigned\\ExcelDemosWithPOI\\WE_2_March\\Charlie_Haddad_Hours_WE_2nd_March_2014.xlsx");
		sourceFileList.put("PayRateDataSheet","E:\\Charlie_Haddad\\demos\\ExcelDemosWithPOI_redesigned\\ExcelDemosWithPOI\\WE_2_March\\PayRate_Correct.xlsx");
		sourceFileList.put("TimeSheetDataSheet","E:\\Charlie_Haddad\\demos\\ExcelDemosWithPOI_redesigned\\ExcelDemosWithPOI\\WE_2_March\\TimeSheet_Correct.xlsx");
		
		//sourceFileList.put("MissedShiftData","D:\\Personal\\ExcelDemos_Redisgned\\ExcelDemosWithPOI\\TimeSheet_Correct_Latest.xlsx");
		
		String destFileName = "E:\\Charlie_Haddad\\demos\\ExcelDemosWithPOI_redesigned\\ExcelDemosWithPOI\\WE_2_March\\Charlie_Haddad_Hours_WE_2nd_March_2014_pavan.xlsx";
		
		
		
		
		
		IReader reader = new ExcelReader();
		
		Model model = reader.read(sourceFileList);
		
		Operations opt = new Operations();
		model = opt.execute(model);
		
		IWriter writer = new MasterExcelWriter();
		writer.write(model,destFileName);
		
		
	}
}
