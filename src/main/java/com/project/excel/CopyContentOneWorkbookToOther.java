package com.project.excel;

import java.awt.Font;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class CopyContentOneWorkbookToOther {

	public static void main(String[] args) throws IOException 
	{
		
		File inputFile=new File("C:\\Test\\CW_test.xlsx");
		FileInputStream fis = new FileInputStream(inputFile);
		XSSFWorkbook inputWorkbook= new XSSFWorkbook(fis);
		int inputSheetCount=inputWorkbook.getNumberOfSheets();
		System.out.println("Input sheetCount: "+inputSheetCount);
		
		
		File outputFile=new File("C:\\Test\\CW_test_output.xlsx");
		FileOutputStream fos=new FileOutputStream(outputFile);
		 
		
		XSSFWorkbook outputWorkbook=new XSSFWorkbook();
		
		
		
		
		for(int i=0;i<inputSheetCount;i++) 
                { 
                  XSSFSheet inputSheet=inputWorkbook.getSheetAt(i); 
                  String inputSheetName=inputWorkbook.getSheetName(i); 
                  XSSFSheet outputSheet=outputWorkbook.createSheet(inputSheetName); 
                 
                                  
                                   
              
                  
                  
                 // Create and call method to copy the sheet and content in new workbook. 
                 copySheet(inputSheet,outputSheet); 
             
                 // Adding content to the Excel sheet
         		addingContent(outputSheet);	
                
         		// Removing Row 
         	//	removeRow(outputSheet, 5);	
         	
         		
         	//	CopyRow(outputSheet, 4);
         		
         		         		
         	//	shiftColumns(headingRow1, 2, 3,outputSheet);
         		
         shiftCells(6, 5 , 5, outputSheet);
         	       		
                }
		
		
		
		
       // Write all the sheets in the new Workbook(testData_Copy.xlsx) using FileOutStream Object
          outputWorkbook.write(fos); 
             
          fos.close(); 
          }

    
	public static void addingContent(XSSFSheet outputSheet)
	{
					
	//	outputSheet.DeleteCells(outputSheet., DeleteMode.EntireRow);
		
		XSSFRow headingRow=outputSheet.createRow(0);
		headingRow.createCell(0).setCellValue("GBS NewBusiness Services");
				
		XSSFRow Product=outputSheet.createRow(1);
		Product.createCell(0).setCellValue("Product");
		
		XSSFRow STATES=outputSheet.createRow(2);
		STATES.createCell(0).setCellValue("STATES");
		STATES.createCell(1).setCellValue("CW");
		
		// Step #6 and Step #7 : Create new Cell and add content to each cell
		XSSFRow Resource=outputSheet.createRow(3);
		Resource.createCell(0).setCellValue("Resource URL");
	
	
		XSSFRow row_5_content=outputSheet.createRow(4);
		row_5_content.createCell(0).setCellValue("Sr No.");
		row_5_content.createCell(1).setCellValue("Test Case ID");
		row_5_content.createCell(2).setCellValue("Test Case Description");
		row_5_content.createCell(3).setCellValue("Service Name");
		row_5_content.createCell(4).setCellValue("Version");
		row_5_content.createCell(5).setCellValue("Product");
		row_5_content.createCell(6).setCellValue("State");
		row_5_content.createCell(7).setCellValue("Line");
		row_5_content.createCell(8).setCellValue("Company");
		row_5_content.createCell(9).setCellValue("Transaction Type");
		row_5_content.createCell(10).setCellValue("Role");
		row_5_content.createCell(11).setCellValue("Channel");
		row_5_content.createCell(12).setCellValue("User/Agent");
		row_5_content.createCell(13).setCellValue("Policy Number");
		row_5_content.createCell(14).setCellValue("Producer Number");
		row_5_content.createCell(15).setCellValue("Transaction Details");
		row_5_content.createCell(16).setCellValue("Test Case Scenario");
		row_5_content.createCell(17).setCellValue("Test Case Input(s)");
		row_5_content.createCell(18).setCellValue("Test Case Expectation(s)");
		row_5_content.createCell(19).setCellValue("DNG Rule(s)");
		row_5_content.createCell(20).setCellValue("Created By");
		row_5_content.createCell(21).setCellValue("Created Date");
		row_5_content.createCell(22).setCellValue("Reviewd By");
		row_5_content.createCell(23).setCellValue("Reviewed Date");
		row_5_content.createCell(24).setCellValue("Comments");
		row_5_content.createCell(25).setCellValue("Is Test Case executed? (Y/N)");
		row_5_content.createCell(26).setCellValue("Test result(Pass/Fail)");
		row_5_content.createCell(27).setCellValue("Defect Number if Failed	Is Fixed?(Y/N)");
		row_5_content.createCell(28).setCellValue("Fix Iteration number");
		row_5_content.createCell(29).setCellValue("Is Configured DB?");
		row_5_content.createCell(30).setCellValue("QA Review Comments");
		row_5_content.createCell(31).setCellValue("Proxy Error messages");
		row_5_content.createCell(32).setCellValue("State Applicable");
		
		
	}

	public static void removeRow(XSSFSheet outputSheet, int rowIndex) {
	    int lastRowNum = outputSheet.getLastRowNum();
	    if (rowIndex >= 0 && rowIndex < lastRowNum) {
	    	outputSheet.shiftRows(rowIndex + 1, lastRowNum, -1);
	    }
	    if (rowIndex == lastRowNum) {
	        Row removingRow = outputSheet.getRow(rowIndex);
	        if (removingRow != null) {
	        	outputSheet.removeRow(removingRow);
	        }
	    }
	}
    
	public static void shiftCells(int startCellNo, int endCellNo, int shiftRowsBy, XSSFSheet outputSheet){
	int lastRowNum = outputSheet.getLastRowNum();
    System.out.println(lastRowNum); //=7
    //      int rowNumAfterAdding = lastRowNum+shiftRowsBy;
    for(int rowNum=lastRowNum;rowNum>0;rowNum--){
        Row rowNew;
        if(outputSheet.getRow((int)rowNum+shiftRowsBy)==null){
            rowNew = outputSheet.createRow((int)rowNum+shiftRowsBy);
        }
        rowNew = outputSheet.getRow((int)rowNum+shiftRowsBy);
        Row rowOld = outputSheet.getRow(rowNum);
        System.out.println("RowNew is "+rowNum+" and Row Old is "+(int)(rowNum+shiftRowsBy));
        System.out.println("startCellNo = "+startCellNo+" endCellNo = "+endCellNo+" shiftRowBy = "+shiftRowsBy);
        for(int cellNo=startCellNo; cellNo<=endCellNo;cellNo++){
            rowNew.createCell(cellNo).setCellValue(rowOld.getCell(cellNo).getStringCellValue().toString());
            rowOld.getCell(cellNo).setCellValue(rowOld.getCell(0).getStringCellValue().toString());
            System.out.println("working on " +cellNo);
            
            
        }
    }       
}
	
	public static void shiftColumns(XSSFRow row, int startingIndex, int shiftCount,XSSFSheet outputSheet) {
	    for (int i = row.getPhysicalNumberOfCells()-1;i>=startingIndex;i--){
	        Cell oldCell = row.getCell(i);
	        Cell newCell = row.createCell(i + shiftCount, oldCell.getCellTypeEnum());
	        cloneCellValue(oldCell,newCell);
	       
	     }
	}

	public static void cloneCellValue(Cell oldCell, Cell newCell) { //TODO test it
	    switch (oldCell.getCellTypeEnum()) {
	        case STRING:
	            newCell.setCellValue(oldCell.getStringCellValue());
	            break;
	        case NUMERIC:
	            newCell.setCellValue(oldCell.getNumericCellValue());
	            break;
	        case BOOLEAN:
	            newCell.setCellValue(oldCell.getBooleanCellValue());
	            break;
	        case FORMULA:
	            newCell.setCellFormula(oldCell.getCellFormula());
	            break;
	        case ERROR:
	            newCell.setCellErrorValue(oldCell.getErrorCellValue());
	        case BLANK:
	        case _NONE:
	            break;
	    }
	}
	
	
	
	
	public static void copySheet(XSSFSheet inputSheet,XSSFSheet outputSheet) 
           { 
              int rowCount=inputSheet.getLastRowNum(); 
              System.out.println(rowCount+" rows in inputsheet "+inputSheet.getSheetName()); 
               
                int currentRowIndex=0; 
                int outcurrentRowIndex=5;
                
                if(rowCount>0)
		{
			Iterator rowIterator=inputSheet.iterator();
			while(rowIterator.hasNext())
			{
				int currentCellIndex=0;
			Iterator cellIterator=((Row) rowIterator.next()).cellIterator();
				
		//		Iterator cellIterator=rowIterator.next().cellIterator();
				
				while(cellIterator.hasNext())
				{
				// Step #5-8 : Creating new Row, Cell and Input value in the newly created sheet. 
					String cellData=cellIterator.next().toString();
					if(currentCellIndex==0)
						outputSheet.createRow(outcurrentRowIndex).createCell(currentCellIndex).setCellValue(cellData);
					else
						outputSheet.getRow(outcurrentRowIndex).createCell(currentCellIndex).setCellValue(cellData);
					
					currentCellIndex++;
				}
				outcurrentRowIndex++;
			}
			System.out.println((currentRowIndex-1)+" rows added to outputsheet "+outputSheet.getSheetName());
			System.out.println();
		}
	}
}
