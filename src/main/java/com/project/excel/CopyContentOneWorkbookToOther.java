package com.project.excel;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class CopyContentOneWorkbookToOther {

	public static void main(String[] args) throws IOException 
	{
		// Step #1 : Locate path and file of input excel.
		File inputFile=new File("C:\\Users\\vikas.singh06\\Desktop\\Task\\testData.xlsx");
		FileInputStream fis = new FileInputStream(inputFile);
		XSSFWorkbook inputWorkbook= new XSSFWorkbook(fis);
		int inputSheetCount=inputWorkbook.getNumberOfSheets();
		System.out.println("Input sheetCount: "+inputSheetCount);
		
		// Step #2 : Locate path and file of output excel.
		File outputFile=new File("C:\\Users\\vikas.singh06\\Desktop\\Task\\testData_output.xlsx");
		FileOutputStream fos=new FileOutputStream(outputFile);
		 
		// Step #3 : Creating workbook for output excel file.
		XSSFWorkbook outputWorkbook=new XSSFWorkbook();
		
		// Step #4 : Creating sheets with the same name as appearing in input file.
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
         		removeRow(outputSheet, 1);	
                
                }
		
       // Write all the sheets in the new Workbook(testData_Copy.xlsx) using FileOutStream Object
          outputWorkbook.write(fos); 
             
          fos.close(); 
          }

    
	public static void addingContent(XSSFSheet outputSheet)
	{
		// Step #5 : Create new Row in the sheet
		XSSFRow headingRow=outputSheet.createRow(0);
		// Step #6 and Step #7 : Create new Cell and add content to each cell
		headingRow.createCell(0).setCellValue("S.No");
		headingRow.createCell(1).setCellValue("Website Name");
		headingRow.createCell(2).setCellValue("URL");
		
		XSSFRow row_1_content=outputSheet.createRow(1);
		row_1_content.createCell(0).setCellValue("1");
		row_1_content.createCell(1).setCellValue("Google");
		row_1_content.createCell(2).setCellValue("www.google.com");
			
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
        
        
		
	
	public static void copySheet(XSSFSheet inputSheet,XSSFSheet outputSheet) 
           { 
              int rowCount=inputSheet.getLastRowNum(); 
              System.out.println(rowCount+" rows in inputsheet "+inputSheet.getSheetName()); 
               
                int currentRowIndex=0; if(rowCount>0)
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
						outputSheet.createRow(currentRowIndex).createCell(currentCellIndex).setCellValue(cellData);
					else
						outputSheet.getRow(currentRowIndex).createCell(currentCellIndex).setCellValue(cellData);
					
					currentCellIndex++;
				}
				currentRowIndex++;
			}
			System.out.println((currentRowIndex-1)+" rows added to outputsheet "+outputSheet.getSheetName());
			System.out.println();
		}
	}
}
