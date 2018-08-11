# DataBasedemo

commons-collections4-4.1
poi-3.17
poi-ooxml-3.17
poi-ooxml-schemas-3.17
xmlbeans-2.6.0

shiftCells(5, 5, 5,  outputSheet);

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
