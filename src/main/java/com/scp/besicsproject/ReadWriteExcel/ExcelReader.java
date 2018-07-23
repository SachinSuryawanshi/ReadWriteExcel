package com.scp.besicsproject.ReadWriteExcel;

import java.io.IOException;
import java.net.URL;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelReader {
    

    public static void main(String[] args)  {
    	ClassLoader classloader =
    			   org.apache.poi.poifs.filesystem.POIFSFileSystem.class.getClassLoader();
    			URL res = classloader.getResource(
    			         "org/apache/poi/poifs/filesystem/POIFSFileSystem.class");
    			String path = res.getPath();
    			System.out.println("Core POI came from " + path);
		
    	XSSFWorkbook workbook;
		try {
			workbook = new XSSFWorkbook("C:\\Users\\Sachin\\Desktop\\Book1.xlsx");
			XSSFSheet sheet = workbook.getSheet("data");
			Iterator<Row> rows = sheet.rowIterator();
			
			while(rows.hasNext()){
					Row row = rows.next();
					Iterator<Cell> cells = row.cellIterator();
							while(cells.hasNext()){
								Cell cell = cells.next();
								System.out.print("\t "+cell.getStringCellValue());
							}
							System.out.println("\n");
			}
			
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
    	
		
    
    
    }
    }
/*
        // Creating a Workbook from an Excel file (.xls or .xlsx)
        Workbook workbook = WorkbookFactory.create(new File(SAMPLE_XLSX_FILE_PATH));

        // Retrieving the number of sheets in the Workbook
        System.out.println("Workbook has " + workbook.getNumberOfSheets() + " Sheets : ");

        
           =============================================================
           Iterating over all the sheets in the workbook (Multiple ways)
           =============================================================
        

        // 1. You can obtain a sheetIterator and iterate over it
        Iterator<Sheet> sheetIterator = workbook.sheetIterator();
        System.out.println("Retrieving Sheets using Iterator");
        while (sheetIterator.hasNext()) {
            Sheet sheet = sheetIterator.next();
            System.out.println("=> " + sheet.getSheetName());
        }*/