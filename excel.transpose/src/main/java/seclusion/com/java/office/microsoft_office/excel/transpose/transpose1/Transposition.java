package seclusion.com.java.office.microsoft_office.excel.transpose.transpose1;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class Transposition {

	public static final String SAMPLE_XLSX_FILE_PATH = "D:\\Utilisateurs\\v-97341M\\Mes documents\\Mes fichiers reçus\\20170724.xlsm";
	public static final String SAMPLE_SHEET_NAME = "Profil";

	public Transposition() {

		
		//https://www.callicoder.com/java-read-excel-file-apache-poi/
		
		try {
			// Creating a Workbook from an Excel file (.xls or .xlsx)
			Workbook workbook = WorkbookFactory.create(new File(SAMPLE_XLSX_FILE_PATH));

			// Retrieving the number of sheets in the Workbook
			System.out.println("Workbook has " + workbook.getNumberOfSheets() + " Sheets : ");

			/*
			 * ============================================================= Iterating over
			 * all the sheets in the workbook (Multiple ways)
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

			transpose(workbook, sheetToWork.getSheetName(), Boolean.FALSE);

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

	public static void transpose(Workbook wb, String sheetName, boolean replaceOriginalSheet) {
		Sheet sheet = wb.getSheet(sheetName);

		AreaReference areaRef = new AreaReference(sheet);
		areaRef.getLastRowAndLastColumn();

		// Pair<Integer, Integer> lastRowColumn = getLastRowAndLastColumn(sheet);
		int lastRow = (int) areaRef.getLastRow(); // lastRowColumn.getFirst();
		int lastColumn = (int) areaRef.getLastColumn(); // lastRowColumn.getSecond();

		System.out.println("Sheet: " + sheet.getSheetName() + "; has " + (lastRow + 1) + " rows and " + lastColumn
				+ " columns, transposing ...");

		// LOG.debug("Sheet {} has {} rows and {} columns, transposing ...", new
		// Object[] {sheet.getSheetName(), 1+lastRow, lastColumn});

		List<CellModel> allCells = new ArrayList<CellModel>();
		for (int rowNum = 0; rowNum <= lastRow; rowNum++) {
			Row row = sheet.getRow(rowNum);
			if (row == null) {
				continue;
			}
			for (int columnNum = 0; columnNum < lastColumn; columnNum++) {
				Cell cell = row.getCell(columnNum);
				allCells.add(new CellModel(cell));
			}
		}

		System.out.println("Read " + allCells.size() + " cells ... transposing them");
		// LOG.debug("Read {} cells ... transposing them", allCells.size());

		Sheet tSheet = wb.createSheet(sheet.getSheetName() + "_transposed");

		for (CellModel cm : allCells) {
			if (cm.isBlank()) {
				continue;
			}

			int tRow = cm.getColNum();
			int tColumn = cm.getRowNum();

			Row row = tSheet.getRow(tRow);
			if (row == null) {
				row = tSheet.createRow(tRow);
			}

			Cell cell = row.createCell(tColumn);
			cm.insertInto(cell);
		}

		for (int i = 0; i < (lastRow + 1); i++) {
			tSheet.autoSizeColumn(i);
			
		}
		
		
		
		AreaReference areaRefFinal = new AreaReference(tSheet);
		areaRefFinal.getLastRowAndLastColumn();
		lastRow = (int) areaRefFinal.getLastRow(); // lastRowColumn.getFirst();
		lastColumn = (int) areaRefFinal.getLastColumn(); // lastRowColumn.getSecond();

		System.out.println("Transposing done. Sheet: " + tSheet.getSheetName() + "; has " + (lastRow + 1) + " rows and "
				+ lastColumn + " columns.");

		// LOG.debug("Transposing done. {} now has {} rows and {} columns.", new
		// Object[] {tSheet.getSheetName(), 1+lastRow, lastColumn});

		/* Write output to file */

		try {
			FileOutputStream output_file = new FileOutputStream(new File("D:/POI_XLS_Pivot_Example.xlsx"));
			wb.write(output_file);// write excel document to output stream
			output_file.close(); // close the file

		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} // create XLSX file
		catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

		/*
		 * if (replaceOriginalSheet) { int pos = wb.getSheetIndex(sheet);
		 * wb.removeSheetAt(pos); wb.setSheetOrder(tSheet.getSheetName(), pos); }
		 */
	}

}
