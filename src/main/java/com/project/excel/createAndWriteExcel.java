package com.project.excel;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class createAndWriteExcel 
{
	public static void main(String[] args) throws IOException 
	{
		String filePath="output_Excel//testResult.xlsx";
		String sheetName="testData";
		createExcel(filePath,sheetName);
	}
	
	public static void createExcel(String filePath,String sheetName) throws IOException
	{
		// Step #1 : Specify file name and location
		File file=new File(filePath);
		
		// Step #2 : load the excel file in FileOutputStream object
		FileOutputStream fos=new FileOutputStream(file); 
		
		// Step #3 : Create new Workbook
		XSSFWorkbook workbook =new XSSFWorkbook(); 
		
		// Step #4 : Create new sheet
		XSSFSheet sheet=workbook.createSheet(sheetName);
			
			// Adding content to the Excel sheet
			addingContent(sheet);
			
			// Write and update Excel File
			writeAndUpdateExcel(workbook, fos);
	}

	public static void addingContent(XSSFSheet sheet)
	{
		// Step #5 : Create new Row in the sheet
		XSSFRow headingRow=sheet.createRow(0);
		// Step #6 and Step #7 : Create new Cell and add content to each cell
		headingRow.createCell(0).setCellValue("S.No");
		headingRow.createCell(1).setCellValue("Website Name");
		headingRow.createCell(2).setCellValue("URL");
		
		// Step #8 : Creating/Adding multiple rows and cells
		XSSFRow row_1_content=sheet.createRow(1);
		row_1_content.createCell(0).setCellValue("1");
		row_1_content.createCell(1).setCellValue("Google");
		row_1_content.createCell(2).setCellValue("www.google.com");
		
		XSSFRow row_2_content=sheet.createRow(2);
		row_2_content.createCell(0).setCellValue("2");
		row_2_content.createCell(1).setCellValue("Facebook");
		row_2_content.createCell(2).setCellValue("www.facebook.com");
		
		XSSFRow row_3_content=sheet.createRow(3);
		row_3_content.createCell(0).setCellValue("3");
		row_3_content.createCell(1).setCellValue("Twitter");
		row_3_content.createCell(2).setCellValue("https://twitter.com/");
		
		XSSFRow row_4_content=sheet.createRow(4);
		row_4_content.createCell(0).setCellValue("4");
		row_4_content.createCell(1).setCellValue("Guru99");
		row_4_content.createCell(2).setCellValue("https://www.guru99.com/");
	}
	
	public static void writeAndUpdateExcel(XSSFWorkbook workbook, FileOutputStream fos) throws IOException
	{
		// Step #9 : Writing all the data from workbook object to FileOutputStream object
		workbook.write(fos);
		
		// Step #10 : Closing the workbook object.
		workbook.close();
		System.out.println("File created and saved at specified location....");
	}
}