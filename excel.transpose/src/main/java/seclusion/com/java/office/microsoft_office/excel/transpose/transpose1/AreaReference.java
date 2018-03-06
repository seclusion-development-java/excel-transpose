package seclusion.com.java.office.microsoft_office.excel.transpose.transpose1;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

public class AreaReference {
	
	private long lastRow;
	private long lastColumn;
	public long getLastRow() {
		return lastRow;
	}

	public void setLastRow(long lastRow) {
		this.lastRow = lastRow;
	}

	public long getLastColumn() {
		return lastColumn;
	}

	public void setLastColumn(long lastColumn) {
		this.lastColumn = lastColumn;
	}

	private Sheet sheet;
	
	public AreaReference(Sheet _sheet) {
		sheet = _sheet;
	}
	
	public void getLastRowAndLastColumn() {
	    int _lastRow = sheet.getLastRowNum();
	    int _lastColumn = 0;
	    for (Row row : sheet) {
	        if (_lastColumn < row.getLastCellNum()) {
	            _lastColumn = row.getLastCellNum();
	        }
	    }
	    
	    lastRow = _lastRow;
	    lastColumn = _lastColumn;
	    
	}
	
	
	
}
