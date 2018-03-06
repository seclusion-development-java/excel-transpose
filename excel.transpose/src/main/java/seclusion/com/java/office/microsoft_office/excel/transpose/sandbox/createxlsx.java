package seclusion.com.java.office.microsoft_office.excel.transpose.sandbox;

import java.io.FileOutputStream;
import java.io.*;

import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.*;
import org.apache.poi.xssf.usermodel.*;

import seclusion.com.java.office.microsoft_office.excel.transpose.transpose1.AreaReference;

import java.util.*;
public class createxlsx {  
        public static void main(String[] args) throws Exception{
                /* Apache POI Create Pivot Table Example Program */
                /* Step -1: Create a workbook object to start with */
                XSSFWorkbook new_workbook = new XSSFWorkbook(); //create a blank workbook object
                /* Create a worksheet in the workbook. We will name it "Pivot Table Example" */
                XSSFSheet sheet = new_workbook.createSheet("Pivot Table Example");  //create a worksheet with caption score_details
                /* Add some Rows and Columns to explain Pivot Table  */         
                /* Create the Header Row */
                Row row1 = sheet.createRow(0);                
                Cell cell11 = row1.createCell(0);
                cell11.setCellValue("Student");
                Cell cell12 = row1.createCell(1);
                cell12.setCellValue("Subject");
                Cell cell13 = row1.createCell(2);
                cell13.setCellValue("Score");
                /* Row #1 */
                Row row2 = sheet.createRow(1);
                Cell cell21 = row2.createCell(0);
                cell21.setCellValue("Matt");
                Cell cell22 = row2.createCell(1);
                cell22.setCellValue("English");
                Cell cell23 = row2.createCell(2);
                cell23.setCellValue(67);
                /* Row #2 */
                Row row3 = sheet.createRow(2);
                Cell cell31 = row3.createCell(0);
                cell31.setCellValue("Pitt");
                Cell cell32 = row3.createCell(1);
                cell32.setCellValue("English");
                Cell cell33 = row3.createCell(2);
                cell33.setCellValue(90);
                /* Row #3 */
                Row row4 = sheet.createRow(3);
                Cell cell41 = row4.createCell(0);
                cell41.setCellValue("Pitt");
                Cell cell42 = row4.createCell(1);
                cell42.setCellValue("Biology");
                Cell cell43 = row4.createCell(2);
                cell43.setCellValue(90);
                /* Row #4 */
                Row row5 = sheet.createRow(4);
                Cell cell51 = row5.createCell(0);
                cell51.setCellValue("Matt");
                Cell cell52 = row5.createCell(1);
                cell52.setCellValue("Physics");
                Cell cell53 = row5.createCell(2);
                cell53.setCellValue(99);
                /* Define an Area Reference for the Pivot Table */
                org.apache.poi.ss.util.AreaReference a=new org.apache.poi.ss.util.AreaReference("A1:C5",SpreadsheetVersion.EXCEL2007);
                /* Define the starting Cell Reference for the Pivot Table */
                CellReference b=new CellReference("I5");
                /* Create the Pivot Table */
                XSSFPivotTable pivotTable = sheet.createPivotTable(a,b);
                /* First Create Report Filter - We want to filter Pivot Table by Student Name */
                pivotTable.addReportFilter(0);
                 /* Second - Row Labels - Once a student is filtered all subjects to be displayed in pivot table */
                pivotTable.addRowLabel(1);
                /* Define Column Label with Function, Sum of the marks obtained */
                pivotTable.addColumnLabel(DataConsolidateFunction.SUM, 2);                
                /* Write output to file */ 
                FileOutputStream output_file = new FileOutputStream(new File("C:/POI_XLS_Pivot_Example.xlsx")); //create XLSX file
                new_workbook.write(output_file);//write excel document to output stream
                output_file.close(); //close the file
        }
}