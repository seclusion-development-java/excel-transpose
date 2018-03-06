package seclusion.com.java.office.microsoft_office.excel.transpose.sandbox;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFPivotTable;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Transpose {

	public static final String SAMPLE_XLSX_FILE_PATH = "./src/main/resources/sample-xlsx-file.xlsx";
	public static final String SAMPLE_SHEET_NAME = "Employee";
	
	public Transpose() {
		transpose();
	}
	
	private void transpose() {
		
		try {
			// Creating a Workbook from an Excel file (.xls or .xlsx)
			Workbook workbook = WorkbookFactory.create(new File(SAMPLE_XLSX_FILE_PATH));
			
			// Retrieving the number of sheets in the Workbook
			System.out.println("Workbook has " + workbook.getNumberOfSheets() + " Sheets : ");
			
			/*
			 * ============================================================= 
			 * Iterating over all the sheets in the workbook (Multiple ways)
			 * =============================================================
			 */
			
			// 1. You can obtain a sheetIterator and iterate over it
			Iterator<Sheet> sheetIterator = workbook.sheetIterator();
			System.out.println("Retrieving Sheets using Iterator");
			while (sheetIterator.hasNext()) {
				Sheet sheet = sheetIterator.next();
				System.out.println("=> " + sheet.getSheetName());
			}
			
			Sheet sheetToWork = workbook.getSheet(SAMPLE_SHEET_NAME);
			
			
			/*
			 * ================================================================== 
			 * Iterating over all the rows and columns in a Sheet (Multiple ways)
			 * ==================================================================
			 */
			
			// Create a DataFormatter to format and get each cell's value as String
			DataFormatter dataFormatter = new DataFormatter();

			// 1. You can obtain a rowIterator and columnIterator and iterate over them
			System.out.println("\n\nIterating over Rows and Columns using Iterator\n");
			Iterator<Row> rowIterator = sheetToWork.rowIterator();
			while (rowIterator.hasNext()) {
				Row row = rowIterator.next();

				// Now let's iterate over the columns of the current row
				Iterator<Cell> cellIterator = row.cellIterator();

				while (cellIterator.hasNext()) {
					Cell cell = cellIterator.next();
					String cellValue = dataFormatter.formatCellValue(cell);
					System.out.print(cellValue + "\t");
				}
				System.out.println();
			}
			
			
			/*
			 * ================================================================== 
			 * Iterating Prepare to transpose in another sheet
			 * ==================================================================
			 */
			
		     /* Step -1: Create a workbook object to start with */
            XSSFWorkbook new_workbook = new XSSFWorkbook(); //create a blank workbook object
            /* Create a worksheet in the workbook. We will name it "Pivot Table Example" */
            XSSFSheet sheet = new_workbook.createSheet("Pivot Table Example");  //create a worksheet with caption score_details
			
            /* Define an Area Reference for the Pivot Table */
            AreaReference a = new AreaReference("A1:E5",SpreadsheetVersion.EXCEL2007);
            
            /* Define the starting Cell Reference for the Pivot Table */
            CellReference b = new CellReference("I5");
            
            
            XSSFSheet sheet2 = (XSSFSheet) new_workbook.createSheet("pivot");
            
            /* Create the Pivot Table */
            XSSFPivotTable pivotTable = sheet.createPivotTable(a,b);
            
         
            
            /* First Create Report Filter - We want to filter Pivot Table by Student Name */
            //pivotTable.addReportFilter(0);
            
            /* Second - Row Labels - Once a student is filtered all subjects to be displayed in pivot table */
            
            //pivotTable.addRowLabel(1);
            /* Define Column Label with Function, Sum of the marks obtained */
            
            //pivotTable.addColumnLabel(DataConsolidateFunction.SUM, 2);                
            
            /* Write output to file */ 
            FileOutputStream output_file = new FileOutputStream(new File("c:/POI_XLS_Pivot_Example.xlsx")); //create XLSX file
            new_workbook.write(output_file);//write excel document to output stream
            output_file.close(); //close the file
			
			
		
			
			
		} catch (EncryptedDocumentException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (InvalidFormatException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		
	}
	
	

}
