package com.santosh.draft;

import java.awt.Desktop;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Scanner;

import javax.swing.JOptionPane;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.xmlbeans.impl.piccolo.io.FileFormatException;

public class CreateTestCase {

	public static  String fileAddress; 
	public static  String fileAddress1; 
	static public  FileInputStream fis = null; 
	static public  FileOutputStream fileOut =null; 
	static public Workbook workbook = null; 
	static private Sheet sheet = null; 
	static private Row row   =null; 
	static private Cell cell   =null; 
	static private String storyName = null;
	static private String responsibleApp = null;

	//Drop1 Variables
	public int Test_Obj, TD_CTN, TD_BAN, TD_SS, TD_PP_Type, TD_PP,TD_SOC, TD_RD, TD_TIL_STUB, Pre_Cond;


	public CreateTestCase() {
		fileAddress = System.getProperty("user.dir")+"//src//com//santosh//res//Zero_Rating_Test_Plan_v0.1.xlsx";
		fileAddress1 = System.getProperty("user.dir")+"//src//com//santosh//res//Zero_Rating_Test_Plan_v0.1_For_Upload.xlsx";
		Test_Obj = 2;
		Pre_Cond = 3;
		TD_CTN = 6;
		TD_BAN = 7;
		TD_SS = 8;
		TD_PP_Type= 9;
		TD_PP=10;
		TD_SOC=11;
		TD_RD=12;
		TD_TIL_STUB=13;

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
			return cellValue;
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
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
		}
		return false;
	}

	public void getTestCaseFileAddressFromConsole() {
		Scanner reader = new Scanner(System.in);  // Reading from System.in
		System.out.println("Enter a Test Case File Address: ");
		fileAddress = reader.nextLine(); // Scans the next token of the input as an int.

		reader.close();
	}

	public void getTestCaseFileAddressFromPopup() {
		//		JTextArea textArea = new JTextArea();
		//        textArea.setEditable(true);
		//        JScrollPane scrollPane = new JScrollPane(textArea);
		//        scrollPane.requestFocus();
		//        textArea.requestFocusInWindow();
		//        scrollPane.setPreferredSize(new Dimension(800, 600));
		//        JOptionPane.showInputDialog( scrollPane,"Enter Test Case Path ", JOptionPane.PLAIN_MESSAGE);
		//		fileAddress = textArea.getText();


		fileAddress = JOptionPane.showInputDialog(this, "Enter the Test Case Path");


		System.out.println(fileAddress);

	}

	/*
	 * Write the excel file in the given file address
	 * @param fileAddress
	 * @return true if successfully write the file; false otherwise
	 * @author kumsanto
	 * 
	 */
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
			// TODO Auto-generated catch block
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

	public String mergeData(Row row) {
		String testDescription = null;

		testDescription = "Test Objective: "+getCellValueInString(row.getCell(Test_Obj));
		////		+"\n\nTest Data: "
		//		+"\n\tBAN: "+)
		//		+"\n\tSource System:  "+)
		//		+"\n\tPrice Plan Type:  "+)
		//		+");"
		//		+ "
		//Pre Condition
		if(getCellValueInString(row.getCell(Pre_Cond)).trim().length()>0){
			testDescription= testDescription 
					+"\n\tPre Condition:  "+getCellValueInString(row.getCell(Pre_Cond));
		}
		//Test Data CTN
		if(getCellValueInString(row.getCell(TD_CTN)).trim().length()>0){
			testDescription= testDescription 
//					+"\n\tCTN:  "+getCellValueInString(row.getCell(TD_CTN));
			+"\n\t"+getCellValueInString(row.getCell(TD_CTN));
		}
		//Test Data BAN
		if(getCellValueInString(row.getCell(TD_BAN)).trim().length()>0){
			testDescription= testDescription 
					+"\n\tBAN:  "+getCellValueInString(row.getCell(TD_BAN));
		}
		//Test Data Source System
		if( getCellValueInString(row.getCell(TD_SS)).trim().length()>0){
			testDescription= testDescription 
					+"\n\tSource System:  "+ getCellValueInString(row.getCell(TD_SS));
		}
		//Test Data Price Plan
		if(getCellValueInString(row.getCell(TD_PP)).trim().length()>0){
			testDescription= testDescription 
					+"\\n\\tPrice Plan:  "+getCellValueInString(row.getCell(TD_PP));
		}
		//Test Data Price Plan Type
		if(getCellValueInString(row.getCell(TD_PP_Type)).trim().length()>0){
			testDescription= testDescription 
					+"\\n\\tPrice Plan Type:  "+getCellValueInString(row.getCell(TD_PP_Type));
		}
		//Test Data SOC
		if(getCellValueInString(row.getCell(TD_SOC)).trim().length()>0){
			testDescription= testDescription 
					+"\n\tSOC:  "+getCellValueInString(row.getCell(TD_SOC));
		}
		//Test Data Roll Over Days
		if(getCellValueInString(row.getCell(TD_RD)).trim().length()>0){
			testDescription= testDescription 
					+"\n\tRollover Day(s):  "+getCellValueInString(row.getCell(TD_RD));
		}//Test Data TD_TIL_STUB
		if(getCellValueInString(row.getCell(TD_TIL_STUB)).trim().length()>0){
			testDescription= testDescription 
					+"\n\tTIL Stub:  "+getCellValueInString(row.getCell(TD_TIL_STUB));
		}
		return testDescription;
		//		System.out.println(testDescription);
	}

	public String setPath(Row row) {
		String path = null;
		String SheetName = row.getSheet().getSheetName();
		SheetName = SheetName.replace('+', '&');
		SheetName = SheetName.replace(' ', '_');
		path = "VFUK-Mobile\\VBC\\MDN_Drop1\\"+SheetName;
		//		System.out.println(path);
		return path;
	}

	public void removeColumns(Row row, int start, int end) {
		System.out.println("Inside removeColumns");
		for (int i =end;i>start;i-- ) {
			System.out.println("Removing cell "+i);
			row.createCell(i);
			row.removeCell(row.getCell(i));

		}
	}

	public void trimCellsOfSheet(Sheet sheet) {
		for(Row row:sheet) {
			for(Cell cell: row) {
				cell.setCellValue(getCellValueInString(cell).trim());
			}
		}
	}

	public void loadExcelFile(String fileAddress) { 
		CreateTestCase.fileAddress=fileAddress; 
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
		//		Row row = cell.getRow();
		//		Sheet sheet = row.getSheet();
		String sheetName = sheet.getSheetName();
		System.out.println("Inside UpdateQC filed With SheetName as "+sheetName);

		if(sheetName.contains("Voice")||sheetName.contains("SMS")||sheetName.contains("MMS")||sheetName.contains("voice")
				||sheetName.contains("capping")||sheetName.contains("Capping")|| sheetName.contains("PIT")||sheetName.contains("Bundle")||sheetName.contains("SOC")) {
			storyName= "Event Processing";
		} else 	if(sheetName.contains("aggregation")||sheetName.contains("Client")) {
			storyName= "Aggregation";
		}else if(sheetName.contains("Regression")) {
			storyName= "Regression";
		} else	if(sheetName.contains("GUI"))
			storyName= "GUI";

		System.out.println(storyName);
		//		switch (sheetName){
		//		case "Domestic Voice": 
		//		case "Domestic SMS":
		//		case "MMS":
		//		case "International voice":
		//		case "Roaming Voice":
		//		case "Roaming SMS":
		//			storyName= "Event Processing";
		//		case "Bundle aggregation":
		//		case "Traveller aggregation":
		//		case "Web Client":
		//			storyName= "Aggregation";
		//		case "Regression":
		//			storyName= "Regression";
		//		case "TAR GUI":
		//			storyName= "GUI";
		//		default:
		//			storyName="Not Found";
		//		}
		// TD_CTN, TD_BAN, TD_SS, TD_PP_Type, TD_PP,TD_SOC, TD_RD, TD_TIL_STUB, Pre_Cond;
		row.createCell(TD_CTN).setCellValue("VBC"); 			//CRID
		row.createCell(TD_BAN).setCellValue("MDN"); 			//Responsible App
		row.createCell(TD_SS).setCellValue("07-System Test");	//Test Level
		row.createCell(TD_PP_Type).setCellValue("CMD"); 		//Topic ID
		row.createCell(TD_PP).setCellValue("Drop1"); 			//Version
		row.createCell(TD_SOC).setCellValue("Design"); 			//Status
		row.createCell(TD_RD).setCellValue(storyName);			//StoryName
		row.createCell(TD_TIL_STUB).setCellValue("MDN"); 		//Detected In App
		row.createCell(TD_TIL_STUB+1).setCellValue("MDN"); 		//Involved App
		row.createCell(TD_TIL_STUB+2).setCellValue("Step 1"); 	//Test Step No


	}

	public void updateTestNameWithTestCaseId(Row row) {
		String prefix = "TC_";
		String testCaseNumber = getCellValueInString(row.getCell(0))+"_";
		String testCaseName = getCellValueInString(row.getCell(1));
		System.out.println();
		row.getCell(1).setCellValue(prefix+testCaseNumber+testCaseName);
	}

	public static void main(String[] args) {

		CreateTestCase ct = new CreateTestCase();
		//		ct.getTestCaseFileAddressFromConsole();
		//		ct.getTestCaseFileAddressFromPopup();
		ct.loadExcelFile(fileAddress);

		//		System.out.println(ct.mergeData(sheet.getRow(3)));

		for (int sheetNo = 1; sheetNo < workbook.getNumberOfSheets()-2; sheetNo ++) {
			System.out.println(workbook.getSheetName(sheetNo));
			sheet = workbook.getSheetAt(sheetNo);
			ct.trimCellsOfSheet(sheet);
			for(Row row : sheet) {

				if(row.getRowNum()>0) {
					cell = row.getCell(2);

					if(ct.getCellValueInString(cell)!=null) {
						cell.setCellValue(ct.mergeData(row));
						row.getCell(ct.Pre_Cond).setCellValue(ct.setPath(row));
						ct.updateQCField(row);
						ct.updateTestNameWithTestCaseId(row);

					}
				}else {
					row.createCell(ct.TD_CTN).setCellValue("CRID"); 			//CRID
					row.createCell(ct.TD_BAN).setCellValue("Responsible_App"); 			//Responsible App
					row.createCell(ct.TD_SS).setCellValue("Test_Level");	//Test Level
					row.createCell(ct.TD_PP_Type).setCellValue("Topic_Id"); 		//Topic ID
					row.createCell(ct.TD_PP).setCellValue("Version"); 			//Version
					row.createCell(ct.TD_SOC).setCellValue("Status"); 			//Status
					row.createCell(ct.TD_RD).setCellValue("Story_Name");			//StoryName
					row.createCell(ct.TD_TIL_STUB).setCellValue("Detected_In_App"); 		//Detected In App
					row.createCell(ct.TD_TIL_STUB+1).setCellValue("Involved_App"); 			//Involved App
					row.createCell(ct.TD_TIL_STUB+2).setCellValue("Test_Step_No");
					row.getCell(ct.Pre_Cond).setCellValue("Path");
				}
			}
		}
		//		System.out.println("Starting removal");
		//		ct.removeColumns(row,ct.TD_TIL_STUB,ct.TD_CTN);
		ct.writeFile(fileAddress1);
		ct.openFile(fileAddress1);

	}

}
