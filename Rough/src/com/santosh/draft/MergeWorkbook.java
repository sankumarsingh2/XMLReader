package com.santosh.draft;

import java.awt.Desktop;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class MergeWorkbook {

	public File[] listOfExcelFiles;
	public static String xlsFileDirectory = null;
	public Sheet sheet1= null;
	public Workbook finalWorkbook = null;

	//create method that takes workbook as parameter and return array of sheets from workbook say getSheetArray.
	public  Sheet[] getSheetArryFromWorkbook( File file) {
		try {
			Workbook wb= new XSSFWorkbook(new FileInputStream(file));
			int sheetCount = wb.getNumberOfSheets();
			Sheet arrayOfSheet[] = new Sheet[sheetCount];
			for(Sheet sheet:wb) {
				arrayOfSheet[wb.getSheetIndex(sheet)]=sheet;
			}
			return arrayOfSheet;
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

		return null;

	}

	public void createNewAndCopySheet(Workbook wb, Sheet sheet) {
		sheet1 =	wb.createSheet(sheet.getSheetName());
		
	}



	public void mergeSheetsOfMultipleWorkbookInSingleWorkbook(String finalXLSAddress) {

		//Risk: Not sure what if the extension is .XSLX or similar
		try {
			finalWorkbook = new XSSFWorkbook(new FileInputStream(new File (finalXLSAddress)));
			for(File file:listOfExcelFiles) {
				Sheet[] sourceSheet = getSheetArryFromWorkbook(file);
				for(Sheet currentSourcesheet:sourceSheet) {
					createNewAndCopySheet(finalWorkbook, currentSourcesheet);
				}
			}

		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}

	public void getAllFilesInDirectoryFromExtension(String pathOfDirectory, String extension) {
		//		System.out.println("Inside getAllFilesInDirectoryFromExtension with values "+pathOfDirectory +", and "+extension);
		File folder = new File(pathOfDirectory);
		listOfExcelFiles = folder.listFiles();
	}



	/**
	 * create method that takes workbook as parameter and return array of sheets from workbook say getSheetArray.
	 * create method that create a new sheet and iterate through each workbooks
	 * 		First iteration will iterate through different workbook
	 * 			Second iteration will call getSheetArray and iterate through it
	 * 				In this loop each time new sheet created in master workbook and each sheet take reference of sheetArray items
	 * @param args
	 */



	/**
	 * Write the excel file in the given file address
	 * @param fileAddress
	 * @return true if successfully write the file; false otherwise
	 * @author kumsanto
	 * 
	 */
	public boolean writeFile(String fileAddress) {
		try {
			File file = new File (fileAddress);
			file.setWritable(true);
			if (file.canWrite()) {
				FileOutputStream fout = new FileOutputStream(new File (fileAddress));
				finalWorkbook.write(fout);
				fout.close();
			}else
				System.out.println("***Unable to Write in the file... Please close all of its open instances");
			//			file.setWritable(false);

		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}


		return false;
	}

/**
 *	 * Open the file given in the file address, if found
	 * @param fileAddress
	 * @return true if file exist and successfully opened it; false otherwise
	 */
	public boolean openFile(String fileAddress) {
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
		return false;
	}




	public static void main(String[] args) {
		MergeWorkbook mb = new MergeWorkbook();
		xlsFileDirectory = System.getProperty("user.dir")+"/src/com/santosh/res";
		mb.getAllFilesInDirectoryFromExtension(xlsFileDirectory+"/xlsDirectory",".xlsx");
		mb.mergeSheetsOfMultipleWorkbookInSingleWorkbook(xlsFileDirectory+"/Test0.xlsx");
		mb.writeFile(xlsFileDirectory+"/Test0.xlsx");
		mb.openFile(xlsFileDirectory+"/Test0.xlsx");
	}

}
