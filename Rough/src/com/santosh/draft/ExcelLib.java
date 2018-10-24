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
import org.apache.poi.ss.usermodel.charts.AxisCrosses;
import org.apache.poi.ss.usermodel.charts.AxisPosition;
import org.apache.poi.ss.usermodel.charts.ChartAxis;
import org.apache.poi.ss.usermodel.charts.ChartDataSource;
import org.apache.poi.ss.usermodel.charts.ChartLegend;
import org.apache.poi.ss.usermodel.charts.DataSources;
import org.apache.poi.ss.usermodel.charts.LegendPosition;
import org.apache.poi.ss.usermodel.charts.LineChartData;
import org.apache.poi.ss.usermodel.charts.LineChartSeries;
import org.apache.poi.ss.usermodel.charts.ValueAxis;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFChart;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFTable;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.xmlbeans.impl.piccolo.io.FileFormatException;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTBoolean;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTLineSer;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTPlotArea;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTTable;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTTableColumn;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTTableColumns;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTTableStyleInfo;

public class ExcelLib {
	public static  String fileAddress; 
	public static  String fileAddress1; 
	static public  FileInputStream fis = null; 
	static public  FileOutputStream fileOut =null; 
	static public Workbook workbook = null; 
	static private Sheet sheet = null; 
	static private Row row   =null; 
	
	
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
				sheet = workbook.getSheetAt(0); 
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

	public boolean checkBlankRow(Row row) {
		if(getCellValueInString(row.getCell(0)).trim().length()==0){
			return false;
		}
		return true;
	}

	public void printSheetNames() {
		for(int sheetNo = 1; sheetNo < workbook.getNumberOfSheets()-2; sheetNo ++) {
			System.out.print("Sheet Name: "+workbook.getSheetAt(sheetNo).getSheetName());
			System.out.println(" - "+workbook.getSheetAt(sheetNo).getLastRowNum());
		}
	}
	
	

	public void D_createLineChart(XSSFSheet sheet1, XSSFSheet sheet2, int NUM_OF_ROWS, int NUM_OF_COLUMNS) {
		

		XSSFDrawing drawing =sheet2.createDrawingPatriarch();
		XSSFClientAnchor anchor = drawing.createAnchor(0, 0, 0, 0, NUM_OF_COLUMNS + 2, 3, NUM_OF_COLUMNS + 15, 20);

		XSSFChart chart = drawing.createChart(anchor);
		ChartLegend legend = chart.getOrCreateLegend();
		legend.setPosition(LegendPosition.RIGHT);

		LineChartData data = chart.getChartDataFactory().createLineChartData();

		ChartAxis bottomAxis = chart.getChartAxisFactory().createCategoryAxis(AxisPosition.BOTTOM);
		ValueAxis leftAxis = chart.getChartAxisFactory().createValueAxis(AxisPosition.LEFT);
		leftAxis.setCrosses(AxisCrosses.AUTO_ZERO);

		ChartDataSource<Number> xs = DataSources.fromNumericCellRange(sheet1, new CellRangeAddress(0, NUM_OF_ROWS - 1, 0, 0));
		ChartDataSource<Number> ys1 = DataSources.fromNumericCellRange(sheet1, new CellRangeAddress(0, NUM_OF_ROWS - 1, 1, 1));
		ChartDataSource<Number> ys2 = DataSources.fromNumericCellRange(sheet1, new CellRangeAddress(0, NUM_OF_ROWS - 1, 2, 2));
		ChartDataSource<Number> ys3 = DataSources.fromNumericCellRange(sheet1, new CellRangeAddress(0, NUM_OF_ROWS - 1, 3, 3));
		ChartDataSource<Number> ys4 = DataSources.fromNumericCellRange(sheet1, new CellRangeAddress(0, NUM_OF_ROWS - 1, 4, 4));

		LineChartSeries series1 = data.addSeries(xs, ys1);
		series1.setTitle("Value1");
		LineChartSeries series2 = data.addSeries(xs, ys2);
		series2.setTitle("Value2");
		LineChartSeries series3 = data.addSeries(xs, ys3);
		series3.setTitle("Value3");
		LineChartSeries series4 = data.addSeries(xs, ys4);
		series4.setTitle("Value4");

		chart.plot(data, bottomAxis, leftAxis);

		XSSFChart xssfChart = (XSSFChart) chart;
		CTPlotArea plotArea = xssfChart.getCTChart().getPlotArea();
		plotArea.getLineChartArray()[0].getSmooth();
		CTBoolean ctBool = CTBoolean.Factory.newInstance();
		ctBool.setVal(false);
		plotArea.getLineChartArray()[0].setSmooth(ctBool);
		for (CTLineSer ser : plotArea.getLineChartArray()[0].getSerArray()) {
			ser.setSmooth(ctBool);

		}

	}

	public void setRowHeight(Sheet sheet, int rowHeightInPixel) {
		for(Row row:sheet)
		row.setHeight((short) rowHeightInPixel);
		}
	
	public void D_convertToTable(XSSFSheet sheet ){
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

}
