package seclusion.com.java.office.microsoft_office.excel.transpose.transpose1;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.RichTextString;

public class CellModel {

	  private int rowNum = -1;
	    private int colNum = -1;
	    private CellStyle cellStyle;
	    private int cellType = -1;
	    private Object cellValue;

	    public CellModel(Cell cell) {
	        if (cell != null) {
	            this.rowNum = cell.getRowIndex();
	            this.colNum = cell.getColumnIndex();
	            this.cellStyle = cell.getCellStyle();
	            this.cellType = cell.getCellType();
	            switch (this.cellType) {
	                case Cell.CELL_TYPE_BLANK:
	                    break;
	                case Cell.CELL_TYPE_BOOLEAN:
	                    cellValue = cell.getBooleanCellValue();
	                    break;
	                case Cell.CELL_TYPE_ERROR:
	                    cellValue = cell.getErrorCellValue();
	                    break;
	                case Cell.CELL_TYPE_FORMULA:
	                    cellValue = cell.getCellFormula();
	                    break;
	                case Cell.CELL_TYPE_NUMERIC:
	                    cellValue = cell.getNumericCellValue();
	                    break;
	                case Cell.CELL_TYPE_STRING:
	                    cellValue = cell.getRichStringCellValue();
	                    break;
	            }
	        }
	    }

	    public boolean isBlank() {
	        return this.cellType == -1 && this.rowNum == -1 && this.colNum == -1;
	    }

	    public void insertInto(Cell cell) {
	        if (isBlank()) {
	            return;
	        }

	        cell.setCellStyle(this.cellStyle);
	        cell.setCellType(this.cellType);
	        switch (this.cellType) {
	            case Cell.CELL_TYPE_BLANK:
	                break;
	            case Cell.CELL_TYPE_BOOLEAN:
	                cell.setCellValue((boolean) this.cellValue);
	                break;
	            case Cell.CELL_TYPE_ERROR:
	                cell.setCellErrorValue((byte) this.cellValue);
	                break;
	            case Cell.CELL_TYPE_FORMULA:
	                cell.setCellFormula((String) this.cellValue);
	                break;
	            case Cell.CELL_TYPE_NUMERIC:
	                cell.setCellValue((double) this.cellValue);
	                break;
	            case Cell.CELL_TYPE_STRING:
	                cell.setCellValue((RichTextString) this.cellValue);
	                break;
	        }
	    }

	    public CellStyle getCellStyle() {
	        return cellStyle;
	    }

	    public int getCellType() {
	        return cellType;
	    }

	    public Object getCellValue() {
	        return cellValue;
	    }

	    public int getRowNum() {
	        return rowNum;
	    }

	    public int getColNum() {
	        return colNum;
	    }
	
}
