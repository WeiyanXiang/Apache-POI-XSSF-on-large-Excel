package com.Domain.processExcel;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.Iterator;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


//Below are essential methods we actually used: it works generically from XLS to XLSM and XLSX to XLSM.

//readAndWriteXLSM("Source.xls", "Template.xlsm");
//readAndWriteXLSM("Source.xlsx", "Template.xlsm");

//Alternative methods provided:
//Below are XSSF approaches examples:
//readXSSF("large.xlsx");
//writeXSSF("large.xlsx", 10, 3);
//readAndWriteXSSF("large.xlsx","largeToWrite.xlsx");

//Below are SXSSF writing related approaches
//writeInSXSSF("TTT.xlsm", 40000, 40);
//readAndWriteSXSSF("Source.xlsx","Template.xlsm");
//readAndWriteVersion2("large.xlsx","largeToWrite.xlsx");
//writeInSXSSFManualFlushing("T1.xlsx", 1000000, 10);
//readAndWriteSXSSFManualFlushing("TTT.xlsx","TTT_copy.xlsx");


public class processExcelFiles {
	
	
	public static void readXSSF(String filePath) throws IOException
	{
		
		InputStream ExcelFileToRead = new FileInputStream(filePath);
	    XSSFWorkbook wb = new XSSFWorkbook(ExcelFileToRead); 
	    Sheet sheet = wb.getSheetAt(0);
	    
		Row row; 
		Cell cell;
		
		Iterator rows = sheet.rowIterator();

		while (rows.hasNext())
		{
			row= (Row) rows.next();
			Iterator cells = row.cellIterator();
			while (cells.hasNext())
			{
				cell= (Cell) cells.next();
		
				if (cell.getCellType() == Cell.CELL_TYPE_STRING)
				{
					System.out.print(cell.getStringCellValue()+"\t");
				}
				else if(cell.getCellType() == Cell.CELL_TYPE_NUMERIC)
				{
					System.out.print(cell.getNumericCellValue()+"\t");
				}
			}
			System.out.println();
		}
		System.out.println("Already read all. ");
		System.gc();
		
		
	}
	
	
	public static void writeXSSF(String filePath, int numOfRow, int numOfCol) throws IOException, InvalidFormatException {
		
		String sheetName = "Sheet1";//name of sheet
		Workbook wb = new XSSFWorkbook();
		Sheet sheet = wb.createSheet(sheetName) ;
		
		for (int r=0;r < numOfRow; r++ )
		{
			Row row = sheet.createRow(r);
			for (int c=0;c < numOfCol; c++ )
			{
				Cell cell = row.createCell(c);
				cell.setCellValue(r+"-"+c);
			}	
		}
		
		FileOutputStream fileOut = new FileOutputStream(filePath);
		
		wb.write(fileOut);
		fileOut.flush();
		fileOut.close();
		System.out.println("Written into xlsx file");
		
	}
	
	public static void readAndWriteXSSF(String fileToRead, String fileToWrite) throws IOException{
		InputStream ExcelFileToRead = new FileInputStream(fileToRead);
		XSSFWorkbook  wb = new XSSFWorkbook(ExcelFileToRead);
		XSSFSheet firstSheet = wb.getSheetAt(0);
		
		XSSFRow row; 
		XSSFCell cell;

		Iterator rows = firstSheet.rowIterator();
		while (rows.hasNext())
		{
 
			row=(XSSFRow) rows.next();
			Iterator cells = row.cellIterator();
			while (cells.hasNext())
			{
				cell=(XSSFCell) cells.next();

				if (cell.getCellType() == XSSFCell.CELL_TYPE_STRING)
				{

					cell.setCellValue(cell.getStringCellValue()+" LL");
				}
				else if(cell.getCellType() == XSSFCell.CELL_TYPE_NUMERIC)
				{
					cell.setCellValue(cell.getNumericCellValue()+" LL");
				}
			}
		}
		
		FileOutputStream fileOut = new FileOutputStream(fileToWrite);
		wb.write(fileOut);
		fileOut.flush();
		fileOut.close();
		System.gc();
		System.out.println("Written into xlsx file");
		
	}
	
	public static void writeInSXSSF(String filePath, int numOfRow, int numOfCol) throws IOException {  
		long startTime = System.nanoTime();
        
		String sheetName = "Sheet1";
        
        // keep 100 rows in memory, exceeding rows will be flushed to disk
        SXSSFWorkbook wb = new SXSSFWorkbook(100); 
	    Sheet sheet = wb.createSheet(sheetName);
	    
	    for (int r=0;r < numOfRow; r++ )
		{
        	Row row = sheet.createRow(r);  
			for (int c=0;c < numOfCol; c++ )
			{
				Cell cell = row.createCell(c);  
	            cell.setCellValue(r + "-" + c);  
			}
		}
        
        FileOutputStream out = new FileOutputStream(filePath);
        wb.write(out);
        out.close();
        
        System.out.println("Write is finished.");
        long endTime = System.nanoTime();
        System.out.println("Time used (in second): " + (endTime-startTime)/1000000000);
	} 
	
	public static void readAndWriteSXSSF(String fileToRead, String fileToWrite) throws IOException{
		long startTime = System.nanoTime();
			
		InputStream ExcelFileToRead = new FileInputStream(fileToRead);
	    XSSFWorkbook wb = new XSSFWorkbook(ExcelFileToRead); 
	    Sheet sheet = wb.getSheetAt(0);
	    
		Row row; 
		Cell cell;
		int numOfRow=0, numOfCol=0;
		
		Iterator rows = sheet.rowIterator();
		
		String sheetName = "Sheet1";
        SXSSFWorkbook SXSSF_wb = new SXSSFWorkbook(100); 
	    Sheet SXSSF_sheet = SXSSF_wb.createSheet(sheetName);

		while (rows.hasNext())
		{
			Row SXSSF_row = SXSSF_sheet.createRow(numOfRow);  
			numOfRow++;
			row= (Row) rows.next();
//			System.out.println(SXSSF_row.getSheet() + "!!!");
			Iterator cells = row.cellIterator();
			while (cells.hasNext())
			{
				Cell SXSSF_cell = SXSSF_row.createCell(numOfCol);
				numOfCol++;
				cell= (Cell) cells.next();
				
				if (cell.getCellType() == Cell.CELL_TYPE_STRING)
				{
//					System.out.print(cell.getStringCellValue()+"\t");
					SXSSF_cell.setCellValue(cell.getStringCellValue() + " ACN");

				}
				else if(cell.getCellType() == Cell.CELL_TYPE_NUMERIC)
				{
//					System.out.print(cell.getNumericCellValue()+"\t");
					SXSSF_cell.setCellValue(cell.getNumericCellValue()+ " ACN"); 

				}
			}
			numOfCol=0;
		}
		
        
        FileOutputStream out = new FileOutputStream(fileToWrite);
        SXSSF_wb.write(out);
        out.close();
        System.gc();
        SXSSF_wb.dispose();
		System.out.println("Written into another xlsx file");
		
        long endTime = System.nanoTime();
        System.out.println("Time used (in second): " + (endTime-startTime)/1000000000);
	}
	
	
	public static void readAndWriteVersion2(String fileToRead, String fileToWrite) throws IOException{
		long startTime = System.nanoTime();
		
		
		InputStream ExcelFileToRead = new FileInputStream(fileToRead);
        XSSFWorkbook wb_template = new XSSFWorkbook(ExcelFileToRead); 
	    SXSSFWorkbook wb = new SXSSFWorkbook(wb_template); 
	    wb.setCompressTempFiles(true);
	    
	    SXSSFSheet sheet = wb.getSheetAt(0);

		// keep 100 rows in memory, exceeding rows will be flushed to disk
	    sheet.setRandomAccessWindowSize(100);
	    Row row;
	    Cell cell;
	    
        Iterator rows = sheet.rowIterator();

		while (rows.hasNext())
		{

			row= (Row) rows.next();
			Iterator cells = row.cellIterator();

			while (cells.hasNext())
			{
				cell= (Cell) cells.next();
		
				if (cell.getCellType() == Cell.CELL_TYPE_STRING)
				{
					cell.setCellValue(cell.getStringCellValue() );  
				}
				else if(cell.getCellType() == Cell.CELL_TYPE_NUMERIC)
				{
					cell.setCellValue(cell.getNumericCellValue() );  

				}
			}
		}
		
		
		FileOutputStream fileOut = new FileOutputStream(fileToWrite);
		wb.write(fileOut);
		fileOut.flush();
		fileOut.close();
		System.gc();
		wb.dispose();
		System.out.println("Written into another xlsx file");
	
	
        long endTime = System.nanoTime();
        System.out.println("Time used (in second): " + (endTime-startTime)/1000000000);
	}
	
	public static void writeInSXSSFManualFlushing(String filePath, int numOfRow, int numOfCol) throws IOException{
		long startTime = System.nanoTime();
		
		// turn off auto-flushing and accumulate all rows in memory
		SXSSFWorkbook wb = new SXSSFWorkbook(-1); 
        Sheet sh = wb.createSheet();
        for(int rownum = 0; rownum < numOfRow; rownum++){
            Row row = sh.createRow(rownum);
            for(int cellnum = 0; cellnum < numOfCol; cellnum++){
                Cell cell = row.createCell(cellnum);
                cell.setCellValue(rownum +"-" + cellnum);
            }
           // manually control how rows are flushed to disk 
           if(rownum % 10000 == 0) {
                ((SXSSFSheet)sh).flushRows(); 
           }

        }

        FileOutputStream out = new FileOutputStream(filePath);
        wb.write(out);
        out.close();
        // dispose of temporary files backing this workbook on disk
        wb.dispose();
        
        System.out.println("Written into another xlsx file");
        long endTime = System.nanoTime();
        System.out.println("Time used (in second): " + (endTime-startTime)/1000000000);
	}
	
	
	public static void readAndWriteSXSSFManualFlushing(String fileToRead, String fileToWrite) throws Throwable{
		long startTime = System.nanoTime();
		
		InputStream ExcelFileToRead = new FileInputStream(fileToRead);
		XSSFWorkbook wb = new XSSFWorkbook(ExcelFileToRead); 
	    Sheet sheet = wb.getSheetAt(0);	    
	    
		Row row; 
		Cell cell;
		int numOfRow=0, numOfCol=0;
		
		Iterator rows = sheet.rowIterator();
		
		String sheetName = "Sheet1";
        SXSSFWorkbook SXSSF_wb = new SXSSFWorkbook(-1);
        Sheet SXSSF_sheet = SXSSF_wb.createSheet(sheetName);

		while (rows.hasNext())
		{
			Row SXSSF_row = SXSSF_sheet.createRow(numOfRow);  
			numOfRow++;
			row= (Row) rows.next();

			Iterator cells = row.cellIterator();
			while (cells.hasNext())
			{
				Cell SXSSF_cell = SXSSF_row.createCell(numOfCol);
				numOfCol++;
				cell= (Cell) cells.next();
				
				if (cell.getCellType() == Cell.CELL_TYPE_STRING)
				{
					SXSSF_cell.setCellValue(cell.getStringCellValue() +" ** ");

				}
				else if(cell.getCellType() == Cell.CELL_TYPE_NUMERIC)
				{
					SXSSF_cell.setCellValue(cell.getNumericCellValue() ); 

				}
			}
			numOfCol=0;
			
            if(numOfRow % 10000 == 0) {
            	((SXSSFSheet)SXSSF_sheet).flushRows();
            }
		}
		
        FileOutputStream out = new FileOutputStream(fileToWrite);
        SXSSF_wb.write(out);
        out.close();
        System.gc();
        SXSSF_wb.dispose();
		
		System.out.println("Written into another xlsx file");
        long endTime = System.nanoTime();
        System.out.println("Time used (in second): " + (endTime-startTime)/1000000000);
	}
	
	
	public static void getAndCopySheet(Sheet originalSheet, Sheet sheetToCopy){
		System.out.println("Daily Data row num: " + originalSheet.getPhysicalNumberOfRows());
		
		for (int r=2;r < originalSheet.getPhysicalNumberOfRows(); r++ )
		{
			Row row2 = sheetToCopy.createRow(r);
			Row row = originalSheet.getRow(r);

			for (int c=1;c < row.getPhysicalNumberOfCells(); c++ )
			{
				Cell cell2 = row2.createCell(c);
				Cell cell = row.getCell(c);
				
				if (cell.getCellType() == Cell.CELL_TYPE_STRING)
				{
					cell2.setCellValue(cell.getStringCellValue());
				}
				else if(cell.getCellType() == Cell.CELL_TYPE_NUMERIC)
				{
					cell2.setCellValue(cell.getNumericCellValue());
				}
			}	
		}
	}
	
	public static void readAndWriteXLSM(String fileToRead, String fileToWrite) throws IOException, InvalidFormatException{
		long startTime = System.nanoTime();
		
		InputStream inp = new FileInputStream(fileToRead);
	    Workbook wb = WorkbookFactory.create(inp);
	    Sheet dailySheet = wb.getSheet("Daily_Actuals");
	    Sheet monthlySheet = wb.getSheet("Monthly_Actuals");
	    
//		This is using XSSF read and write.
//		InputStream inp2 = new FileInputStream(fileToWrite);
//	    Workbook wb2 = WorkbookFactory.create(inp2);
//	    Sheet sheet2= wb2.getSheet("Raw Data - Daily");
	    
//		This is using XSSF read but SXSSF write.
	    InputStream inp2 = new FileInputStream(fileToWrite);
	    XSSFWorkbook XSSFWb = (XSSFWorkbook) WorkbookFactory.create(inp2);
	    Sheet dailySheetToWrite = XSSFWb.getSheet("Raw Data - Daily");
	    Sheet monthlySheetToWrite = XSSFWb.getSheet("Raw Data - Monthly");
	    SXSSFWorkbook SXSSFWb = new SXSSFWorkbook(XSSFWb);
		
	    
		getAndCopySheet(dailySheet, dailySheetToWrite);
		getAndCopySheet(monthlySheet, monthlySheetToWrite);
		
		FileOutputStream fileOut2 = new FileOutputStream(fileToWrite);
		SXSSFWb.write(fileOut2);
		fileOut2.flush();
		fileOut2.close();
		SXSSFWb.dispose();
		
		System.out.println("Done processing the file");
		long endTime = System.nanoTime();
        System.out.println("Total time used (in second): " + (endTime-startTime)/1000000000);
	}
	
	public static void main(String args[]) throws Throwable{
		
//		Below are XSSF approaches.
//		readXSSF("large.xlsx");
//		writeXSSF("large.xlsx", 10, 3);
//		readAndWriteXSSF("large.xlsx","largeToWrite.xlsx");

//		Below are SXSSF writing related approaches
//		writeInSXSSF("TTT.xlsm", 40000, 40);
//		readAndWriteSXSSF("Source.xlsx","Template.xlsm");
//		readAndWriteVersion2("large.xlsx","largeToWrite.xlsx");
//		writeInSXSSFManualFlushing("T1.xlsx", 1000000, 10);
//		readAndWriteSXSSFManualFlushing("TTT.xlsx","TTT_copy.xlsx");
		
		// Actuals Report w e.xlsm
//		readAndWriteXLSM("Source.xls", "Template.xlsm");
		readAndWriteXLSM("Source.xlsx", "Template.xlsm");
	}

}

