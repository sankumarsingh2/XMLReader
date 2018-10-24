package com.santosh.draft;

import java.util.HashMap;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

public class Drop2ReceivedFormat {

	public int Test_ID_ColNo, Test_Case_ColNo, Test_Objective_ColNo, Pre_condition_ColNo, Test_Steps_ColNo, Expected_result_ColNo, TD_CTN_ColNo, TD_SS_ColNo, TD_PP_Type_ColNo, TD_PP_ColNo, TD_VAP_ColNo, Result_ColNo;
	public String Test_ID, Test_Case, Test_Objective, Pre_condition, Test_Steps, Expected_result, TD_CTN, TD_SS, TD_PP_Type, TD_PP, TD_VAP, Result;
	
	public Map <String, String> sourceTestCase = new HashMap <>();
	public Map <String, Integer> sourceColumnNo = new HashMap<>();
	public Map <String, String> destTestCase = new HashMap <>();
	public Map <String, Integer> destColumnNo = new HashMap<>();
	
	public void initSourceColVariables(){
		sourceColumnNo.put("TCID_ColNo" , 		1);
		sourceColumnNo.put("TC_ColNo" ,			2);
		sourceColumnNo.put("TObj_ColNo" ,	 	3); 
		sourceColumnNo.put("PC_ColNo" ,			4);
		sourceColumnNo.put("TStep_ColNo" ,		5);
		sourceColumnNo.put("ExpRes_ColNo" ,		6);
		sourceColumnNo.put("TD_CTN_ColNo" , 	7);
		sourceColumnNo.put("TD_SS_ColNo" , 		8);
		sourceColumnNo.put("TD_PPT_ColNo", 		9);
		sourceColumnNo.put("TD_PP_ColNo" , 		10);
		sourceColumnNo.put("TD_VAP_ColNo" , 	11);
		sourceColumnNo.put("Result_ColNo" , 	12);
	}

	public void initDestColVariables(){
		destColumnNo.put("TCID_ColNo" , 			1);
		destColumnNo.put("TestCase_ColNo" ,			2);
		destColumnNo.put("TestDescription_ColNo" ,	3); 
		destColumnNo.put("Path_ColNo" ,				4);
		destColumnNo.put("TestStep_ColNo" ,			5);
		destColumnNo.put("ExpectedResult_ColNo" ,	6);
		destColumnNo.put("CRID_ColNo" , 			7);
		destColumnNo.put("RespposibleApp_ColNo" , 	8);
		destColumnNo.put("TestLevel_ColNo", 		9);
		destColumnNo.put("TopicID_ColNo" , 			10);
		destColumnNo.put("Version_ColNo" , 			11);
		destColumnNo.put("Status_ColNo" , 			12);	
		destColumnNo.put("StoryName_ColNo" , 		13);	
		destColumnNo.put("DetectedInApp_ColNo" , 	14);	
		destColumnNo.put("InvolvedIApp_ColNo" , 	15);	
		destColumnNo.put("TestStepNo_ColNo" , 		16);	
	}
	
	public void initsourceTestCaseMap() {
		sourceTestCase.put(Test_ID, null);
		sourceTestCase.put(Test_Case, null);
		sourceTestCase.put(Test_Objective, null);
		sourceTestCase.put(Pre_condition, null);
		sourceTestCase.put(Test_Steps, null);
		sourceTestCase.put(Expected_result, null);
		sourceTestCase.put(TD_CTN, null);
		sourceTestCase.put(TD_SS, null);
		sourceTestCase.put(TD_PP, null);
		sourceTestCase.put(TD_PP_Type, null);
		sourceTestCase.put(TD_VAP, null);
		sourceTestCase.put(Result, null);
	}
	
	public void initDestTestCaseMap() {
		destTestCase.put("TCID" , 				null);
		destTestCase.put("TestCase" ,			null);
		destTestCase.put("TestDescription" ,	null);
		destTestCase.put("Path" ,				null);
		destTestCase.put("TestStep" ,			null);
		destTestCase.put("ExpectedResult" ,		null);
		destTestCase.put("CRID" , 				"VBC");
		destTestCase.put("RespposibleApp" , 	"Subadmin");
		destTestCase.put("TestLevel", 			"07-System Test");
		destTestCase.put("TopicID" , 			"CMD");
		destTestCase.put("Version" , 			"Drop2");
		destTestCase.put("Status" , 			"Design");	
		destTestCase.put("StoryName" , 			null);	
		destTestCase.put("DetectedInApp" , 		"Subadmin");
		destTestCase.put("InvolvedIApp" , 		"Subadmin");
		destTestCase.put("TestStepNo" , 		null);
	}
	
	public void initColValues() {
		Test_ID = null;
		Test_Case = null;
		Test_Objective = null;
		Pre_condition = null;
		Test_Steps = null;
		Expected_result = null;
		TD_CTN = null;
		TD_SS = null;
		TD_PP_Type = null;
		TD_PP = null;
		TD_VAP = null;
		Result = null;

	}
	
	public Drop2ReceivedFormat() {
		
	}
	
	public void updateDrop2TestCaseFromRow(Row row) {
		Test_ID = getCellValueInString(row.getCell(Test_ID_ColNo));
		Test_Case = getCellValueInString(row.getCell(Test_Case_ColNo));
		Test_Objective = getCellValueInString(row.getCell(Test_Objective_ColNo));
		Pre_condition = getCellValueInString(row.getCell(Pre_condition_ColNo));
		Test_Steps = getCellValueInString(row.getCell(Test_Steps_ColNo));
		Expected_result = getCellValueInString(row.getCell(Expected_result_ColNo));
		TD_CTN = getCellValueInString(row.getCell(TD_CTN_ColNo));
		TD_SS = getCellValueInString(row.getCell(TD_SS_ColNo));
		TD_PP_Type = getCellValueInString(row.getCell(TD_PP_Type_ColNo));
		TD_PP = getCellValueInString(row.getCell(TD_PP_ColNo));
		TD_VAP = getCellValueInString(row.getCell(TD_VAP_ColNo));
		Result = getCellValueInString(row.getCell(Result_ColNo));
	}
	
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

	public static void main(String[] args) {
		// TODO Auto-generated method stub

	}

}
