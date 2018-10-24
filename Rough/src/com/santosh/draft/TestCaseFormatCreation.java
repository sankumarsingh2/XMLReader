package com.santosh.draft;

import java.awt.Desktop;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFTable;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.xmlbeans.impl.piccolo.io.FileFormatException;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTTable;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTTableColumn;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTTableColumns;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTTableStyleInfo;

public class TestCaseFormatCreation {

	public static  String fileAddress; 
	public static  String fileAddress1; 
	static public  FileInputStream fis = null; 
	static public  FileOutputStream fileOut =null; 
	static public Workbook workbook = null; 
	static private Sheet sheet = null; 
	static private Row row   =null; 
	static private Cell cell   =null; 
	static private String storyName = null;
	public String DROP_NUMBER = "DROP3";

	//Common Variables
	public int TCID_ColNo, TC_ColNo, TObj_ColNo, PC_ColNo, TStep_ColNo, ExpRes_ColNo, TD_CTN_ColNo, TD_PPT_ColNo, TD_PP_ColNo ;

	//Drop1 Variables
	public int TD_BAN, TD_SOC, TD_RD, TD_TIL_STUB;

	//Drop 2 
	public int  TD_SS_ColNo, TD_VAP_ColNo, Result_ColNo;

	/**
	 * constructor
	 */
	public TestCaseFormatCreation() {

		//General or common fields
		TCID_ColNo 		=	0;
		TC_ColNo 		=	1;
		TObj_ColNo 		= 	2;
		PC_ColNo 		=	3;
		TStep_ColNo 	=	4;
		ExpRes_ColNo	=	5;
		TD_CTN_ColNo 	= 	6;

		//Drop 1 Specific
		if(DROP_NUMBER.equalsIgnoreCase("Drop1")||DROP_NUMBER.equalsIgnoreCase("Drop3")) {
			if(DROP_NUMBER.equalsIgnoreCase("Drop1")) {
				fileAddress = System.getProperty("user.dir")+"//src//com//santosh//res//TestCaseDrop1.xlsx";
				fileAddress1= System.getProperty("user.dir")+"//src//com//santosh//res//TestCaseDrop1_For_Upload.xlsx";
			}
			if(DROP_NUMBER.equalsIgnoreCase("Drop3")) {
				fileAddress = System.getProperty("user.dir")+"//src//com//santosh//res//TestCaseDrop3.xlsx";
				fileAddress1= System.getProperty("user.dir")+"//src//com//santosh//res//TestCaseDrop3_For_Upload.xlsx";
			}
			TD_BAN 		=	7;
			TD_SS_ColNo = 	8;
			TD_PPT_ColNo= 	9;
			TD_PP_ColNo =	10;
			TD_SOC		=	11;
			TD_RD		=	12;
			TD_TIL_STUB	=	13;
		}

		//Drop 2 Specific
		if(DROP_NUMBER.equalsIgnoreCase("Drop2")) {
			fileAddress = System.getProperty("user.dir")+"//src//com//santosh//res//TestCaseDrop2_temp.xlsx";
			fileAddress1 = System.getProperty("user.dir")+"//src//com//santosh//res//TestCaseDrop2_For_Upload.xlsx";
			TD_SS_ColNo 	= 	7;
			TD_PPT_ColNo	= 	8;
			TD_PP_ColNo 	=	9;
			TD_VAP_ColNo 	= 	10;
			Result_ColNo 	= 	11;
		}

	}

	
		
	/*
	 * get Cell value as String 
	 * Cell type includes Boolean, Numeric, String, Formula, or Blank
	 * @param cell
	 * @return
	 * @author kumsanto
	 */
	public String getCellValueInString(Cell cell) {

		try {
			String cellValue="";
			int cellType = cell.getCellType();

			switch (cellType) {
			case Cell.CELL_TYPE_BOOLEAN:
				cellValue = cell.getBooleanCellValue() +"";
				break;
			case Cell.CELL_TYPE_NUMERIC:
				cell.setCellType(Cell.CELL_TYPE_STRING);
				cellValue = cell.getStringCellValue() + "";
				break;
			case Cell.CELL_TYPE_STRING:
				cellValue = cell.getStringCellValue();
				break;
			case Cell.CELL_TYPE_FORMULA:
				cellValue = cell.getCellFormula();
				break;
			case Cell.CELL_TYPE_BLANK:
				break;

			default:
				break;
			}
			return cellValue.trim();
		}catch(NullPointerException e) {
			System.out.println("Cell is null");
			//			e.printStackTrace();
			return "";
		}
	}

	/**
	 * Open the file given in the file address, if found
	 * @param fileAddress
	 * @return true if file exist and successfully opened it; false otherwise
	 */
	public boolean openFile(String fileAddress) {
		if(isFileExist(fileAddress)) {
			File file = new File (fileAddress);
			//			file.setReadOnly();
			Desktop desktop = Desktop.getDesktop();
			try {
				desktop.open(file);
				return true;
			} catch (IOException e) {
				e.printStackTrace();
			}
		}
		return false;
	}

	public boolean writeFile(String fileAddress) {
		try {
			File file = new File (fileAddress);
			//			file.setWritable(true);
			FileOutputStream fout = new FileOutputStream(new File (fileAddress));
			workbook.write(fout);
			fout.close();
			if (file.canWrite()) {

			}else
				System.out.println("***Unable to Write in the file... Please close all of its open instances");
			//			file.setWritable(false);

		} catch (IOException e) {
			e.printStackTrace();
		}


		return false;
	}

	public boolean isFileExist(String fileAddress) {

		File file = new File(fileAddress);
		if( file.exists() && file.isFile()) {
			//			System.out.println("File found in location "+ fileAddress);
			return true;
		}
		else if(file.isDirectory()) {
			System.out.println("The address "+ fileAddress+" is not a file but a directory");
		}
		else	
			System.out.println("File doesn't exist in location "+ fileAddress);
		return false;
	}

	public String getDataForColumn(String columnName, Cell cell) {
		getCellValueInString(cell);
		if(!getCellValueInString(cell).equalsIgnoreCase("")){
			return "\n"+columnName+":  "+getCellValueInString(cell);
		}
		else 
			return"";
	}

	public String mergeDataDrop(Row row) {
		String testDescription = null;

		if(DROP_NUMBER.equalsIgnoreCase("Drop1")||DROP_NUMBER.equalsIgnoreCase("Drop3")) {
			testDescription = getDataForColumn("Test Objective",row.getCell(TObj_ColNo))
					+getDataForColumn("Pre Condition",row.getCell(PC_ColNo))
					+"\n\nTest Data: "
					+getDataForColumn("CTN",row.getCell(TD_CTN_ColNo))
					+getDataForColumn("BAN",row.getCell(TD_BAN))
					+getDataForColumn("Source System",row.getCell(TD_SS_ColNo))
					+getDataForColumn("Price Plan Type",row.getCell(TD_PPT_ColNo))
					+getDataForColumn("Price Plan",row.getCell(TD_PP_ColNo))
					+getDataForColumn("SOC",row.getCell(TD_SOC))
					+getDataForColumn("Rollover Day(s)",row.getCell(TD_RD))
					+getDataForColumn("TIL Stub",row.getCell(TD_TIL_STUB));

		}

		if(DROP_NUMBER.equalsIgnoreCase("Drop2")) {
			testDescription = getDataForColumn("Test Objective",row.getCell(TObj_ColNo))
					+getDataForColumn("Pre Condition",row.getCell(PC_ColNo))
					+"\n\nTest Data: "
					+getDataForColumn("CTN",row.getCell(TD_CTN_ColNo))
					+getDataForColumn("Source System",row.getCell(TD_SS_ColNo))
					+getDataForColumn("Price Plan Type",row.getCell(TD_PPT_ColNo))
					+getDataForColumn("Price Plan",row.getCell(TD_PP_ColNo))
					+getDataForColumn("Value Added Product",row.getCell(TD_VAP_ColNo));
		}
		return testDescription.trim();
	}

	public String setPath(Row row) {
		String path = null;
		String SheetName = row.getSheet().getSheetName();
		SheetName = SheetName.replace('+', '&');
		SheetName = SheetName.replace(' ', '_');
		path = "VFUK-Mobile\\VBC\\MDN_"+DROP_NUMBER+"\\"+SheetName;
		//		System.out.println(path);
		return path;
	}

	public void trimCellsOfSheet(Sheet sheet) {
		for(Row row:sheet) {
			for(Cell cell: row) {
				cell.setCellValue(getCellValueInString(cell).trim());
			}
		}
	}

	public void loadExcelFile(String fileAddress) { 
		TestCaseFormatCreation.fileAddress=fileAddress; 
		if(isFileExist(fileAddress)) {
			//			setReadOnly(fileAddress);
			try { 
				fis = new FileInputStream(fileAddress); 
				workbook = new XSSFWorkbook(fis); 
				sheet = workbook.getSheetAt(7); 
				row = sheet.getRow(0);
				row.getCell(0);
				fis.close(); 
			} catch (FileFormatException e) { 
				System.out.println("Not an excel file type " + e.getMessage());
				e.printStackTrace(); 
			}catch (Exception e) { 
				e.printStackTrace(); 
			} 
		}
		else
			System.out.println("File Address is not correct.");
	} 

	public void updateQCField(Row row) {
		String sheetName = sheet.getSheetName();
		if(DROP_NUMBER.equalsIgnoreCase("Drop1")||DROP_NUMBER.equalsIgnoreCase("Drop3")) {
			if(sheetName.contains("oice")||sheetName.contains("ating")||sheetName.contains("SMS")||sheetName.contains("MMS")) {
				storyName= "Event Processing";
			}
			if(sheetName.contains("aggregation")||sheetName.contains("Client"))
				storyName= "Aggregation";
			if(sheetName.contains("Regression"))
				storyName= "Regression";
			if(sheetName.contains("GUI"))
				storyName= "GUI";
			if(sheetName.contains("Opt in and out"))
				storyName= "Amend";
			if(sheetName.contains("barring"))
				storyName= "Cancel";
			if(sheetName.contains("data sharer"))
				storyName= "Fallout";
			if(sheetName.contains("Other"))
				storyName= "Tools";
			row.createCell(TD_CTN_ColNo).setCellValue("VBC"); 			//CRID
			row.createCell(TD_BAN).setCellValue("MDN"); 				//Responsible App
			row.createCell(TD_SS_ColNo).setCellValue("07-System Test");	//Test Level
			row.createCell(TD_PPT_ColNo).setCellValue("CMD"); 			//Topic ID
			row.createCell(TD_PP_ColNo).setCellValue(DROP_NUMBER);	 		//Version
			row.createCell(TD_SOC).setCellValue("Design"); 				//Status
			row.createCell(TD_RD).setCellValue(storyName);				//StoryName
			row.createCell(TD_TIL_STUB).setCellValue("MDN");	 		//Detected In App
			row.createCell(TD_TIL_STUB+1).setCellValue("MDN"); 			//Involved App
			row.createCell(TD_TIL_STUB+2).setCellValue("Step 1"); 		//Test Step No
		}
		if(DROP_NUMBER.equalsIgnoreCase("Drop2")) {
			storyName = "Subadmin";

			//		System.out.println(storyName);
			row.createCell(TD_CTN_ColNo).setCellValue("VBC"); 				//CRID
			row.createCell(TD_SS_ColNo).setCellValue("Subadmin"); 			//Responsible App
			row.createCell(TD_PPT_ColNo).setCellValue("07-System Test");	//Test Level
			row.createCell(TD_PP_ColNo).setCellValue("CMD"); 				//Topic ID
			row.createCell(TD_VAP_ColNo).setCellValue(DROP_NUMBER); 			//Version
			row.createCell(TD_VAP_ColNo+1).setCellValue("Design"); 			//Status
			row.createCell(TD_VAP_ColNo+2).setCellValue(storyName);			//StoryName
			row.createCell(TD_VAP_ColNo+3).setCellValue("Subadmin"); 		//Detected In App
			row.createCell(TD_VAP_ColNo+4).setCellValue("Subadmin"); 		//Involved App
			row.createCell(TD_VAP_ColNo+5).setCellValue("Step 1"); 			//Test Step No
		}

	}

	public void updateTestNameWithTestCaseId(Row row) {
		String prefix = "TC_";
		String testCaseNumber = getCellValueInString(row.getCell(0));
		String testCaseName = "_"+getCellValueInString(row.getCell(1));
		//		System.out.println("Test Case Number is "+testCaseNumber);

		if(testCaseNumber.length()==2 ) {
			prefix = "TC_0";
		}else if(testCaseNumber.length()==1 ) {
			prefix = "TC_00";
		}

		row.getCell(1).setCellValue(prefix+testCaseNumber+testCaseName);
	}

	public void createHeader(Sheet sheet) {
		Row headerRow = sheet.getRow(0);
		if(DROP_NUMBER.equalsIgnoreCase("Drop1")||DROP_NUMBER.equalsIgnoreCase("Drop3")) {
			headerRow.createCell(TD_CTN_ColNo).setCellValue("CRID"); 			//CRID
			headerRow.createCell(TD_BAN).setCellValue("Responsible_App"); 			//Responsible App
			headerRow.createCell(TD_SS_ColNo).setCellValue("Test_Level");	//Test Level
			headerRow.createCell(TD_PPT_ColNo).setCellValue("Topic_Id"); 		//Topic ID
			headerRow.createCell(TD_PP_ColNo).setCellValue("Version"); 			//Version
			headerRow.createCell(TD_SOC).setCellValue("Status"); 			//Status
			headerRow.createCell(TD_RD).setCellValue("Story_Name");			//StoryName
			headerRow.createCell(TD_TIL_STUB).setCellValue("Detected_In_App"); 		//Detected In App
			headerRow.createCell(TD_TIL_STUB+1).setCellValue("Involved_App"); 			//Involved App
			headerRow.createCell(TD_TIL_STUB+2).setCellValue("Test_Step_No");
			headerRow.getCell(PC_ColNo).setCellValue("Path");
		}

		if(DROP_NUMBER.equalsIgnoreCase("Drop2")) {
			headerRow.createCell(TD_CTN_ColNo).setCellValue("CRID"); 			//CRID
			headerRow.createCell(TD_SS_ColNo).setCellValue("Responsible_App"); 			//Responsible App
			headerRow.createCell(TD_PPT_ColNo).setCellValue("Test_Level");	//Test Level
			headerRow.createCell(TD_PP_ColNo).setCellValue("Topic_Id"); 		//Topic ID
			headerRow.createCell(TD_VAP_ColNo).setCellValue("Version"); 			//Version
			headerRow.createCell(TD_VAP_ColNo+1).setCellValue("Status"); 			//Status
			headerRow.createCell(TD_VAP_ColNo+2).setCellValue("Story_Name");			//StoryName
			headerRow.createCell(TD_VAP_ColNo+3).setCellValue("Detected_In_App"); 		//Detected In App
			headerRow.createCell(TD_VAP_ColNo+4).setCellValue("Involved_App"); 			//Involved App
			headerRow.createCell(TD_VAP_ColNo+5).setCellValue("Test_Step_No");
			headerRow.getCell(PC_ColNo).setCellValue("Path");
		}
	}

	public boolean checkBlankRow(Row row) {
		if(getCellValueInString(row.getCell(0)).trim().length()==0){
			return false;
		}
		return true;
	}

	public void printSheetNames() {
		for(int sheetNo = 1; sheetNo < workbook.getNumberOfSheets()-2; sheetNo ++) {
			System.out.println(workbook.getSheetAt(sheetNo).getSheetName());
		}
	}

	public void convertToTable(XSSFSheet sheet ){
		int maxRow = sheet.getLastRowNum();
		int maxCol = sheet.getRow(0).getLastCellNum();
		/* Create Table into Existing Worksheet */
		XSSFTable my_table = sheet.createTable();    
		/* get CTTable object*/
		CTTable cttable = my_table.getCTTable();
		/* Define Styles */    
		CTTableStyleInfo table_style = cttable.addNewTableStyleInfo();
		table_style.setName("TableStyleMedium9");           
		/* Define Style Options */
		table_style.setShowColumnStripes(false); //showColumnStripes=0
		table_style.setShowRowStripes(true); //showRowStripes=1    
		/* Define the data range including headers */
		@SuppressWarnings("deprecation")
		AreaReference my_data_range = new AreaReference(new CellReference(0, 0), new CellReference(maxRow,maxCol));    
		/* Set Range to the Table */
		cttable.setRef(my_data_range.formatAsString());
		cttable.setDisplayName("MYTABLE");      /* this is the display name of the table */
		cttable.setName("Test");    /* This maps to "displayName" attribute in &lt;table&gt;, OOXML */            
		cttable.setId(1L); //id attribute against table as long value
		/* Add header columns */               
		CTTableColumns columns = cttable.addNewTableColumns();
		columns.setCount(3L); //define number of columns
		/* Define Header Information for the Table */
		for (int i = 0; i < 3; i++)
		{
			CTTableColumn column = columns.addNewTableColumn();   
			column.setName("Column" + i);      
			column.setId(i+1);
		}   
	}

	public static void main(String[] args) {
		int count = 0;
		TestCaseFormatCreation ct = new TestCaseFormatCreation();
		ct.loadExcelFile(fileAddress);
		for (int sheetNo = 1; sheetNo < workbook.getNumberOfSheets()-2; sheetNo ++) {
			//System.out.println(workbook.getSheetName(sheetNo));
			sheet = workbook.getSheetAt(sheetNo);
			ct.trimCellsOfSheet(sheet);
			for(Row row : sheet) {
				row.setHeight((short) 300);
				if(row.getRowNum()>0) {
					cell = row.getCell(2);
					if(ct.getCellValueInString(cell)!=null&&(ct.getCellValueInString(row.getCell(0)).trim().length()!=0)) {
						cell.setCellValue(ct.mergeDataDrop(row));
						row.getCell(ct.PC_ColNo).setCellValue(ct.setPath(row));
						ct.updateQCField(row);
//						ct.updateTestNameWithTestCaseId(row);
						count++;

					}
				}else {
					ct.createHeader(sheet);
				}
			}
			//		ct.convertToTable((XSSFSheet) sheet);	
		}
		//		System.out.println("Starting removal");
		//		ct.removeColumns(row,ct.TD_TIL_STUB,ct.TD_CTN);
		System.out.println("Total Test Case Count = "+count);
		ct.printSheetNames();
		ct.writeFile(fileAddress1);
		ct.openFile(fileAddress1);

	}

}



