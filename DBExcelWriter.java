package com.sl;
import java.io.FileOutputStream;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class DBExcelWriter {

    private static String dest = "E:\\CITI_PROJECTS\\Projects\\DestinationFile.xlsx";
	private static XSSFWorkbook myWorkBook = new XSSFWorkbook();
	private static XSSFSheet SourceSheet = myWorkBook.createSheet();
	private static XSSFSheet TargetSheet = myWorkBook.createSheet();
	

	private static void SourceexcelLog(int row, int col, String Sourcevalue1) {
	    XSSFRow SourceRow = SourceSheet.getRow(row);

	    /**
	    if (SourceRow == null)
	    	SourceRow = SourceSheet.createRow(row);

	    //XSSFCell SourceCell = SourceRow.createCell(col);
	    XSSFCell SourceCell = SourceRow.getCell(col);
	    
	    if (SourceCell == null)
	       SourceCell = SourceRow.createCell(col);
	    
	    SourceCell.setCellValue(Sourcevalue1);
	
	  **/
	
	    if (SourceRow == null)
	    	SourceRow = SourceSheet.createRow(row);

	    //XSSFCell SourceCell = SourceRow.createCell(col);
	   // XSSFCell SourceCell = SourceRow.getCell(col);
	    
	    //if (SourceCell == null)
	    XSSFCell SourceCell = SourceRow.createCell(col);
	    
	    SourceCell.setCellValue(Sourcevalue1);
	
	}
	
	
	private static void TargetexcelLog(int row, int col, String Targetvalue1) {
	    XSSFRow TargetRow = TargetSheet.getRow(row);

	    if (TargetRow == null)
	    	TargetRow = TargetSheet.createRow(row);
	    
	    XSSFCell TargetCell = TargetRow.getCell(col);

	    //XSSFCell TargetCell = TargetRow.createCell(col);
	    if (TargetCell == null)
	    	TargetCell = TargetRow.createCell(col);
	    
	       TargetCell.setCellValue(Targetvalue1);
	}
	
	

	public static void pushSourceDatatoExcel(int rowcount,int Totalcolumncount,String SourceMismatchValue) {
	    int TotalnumCol = Totalcolumncount; // assume 10 cols

	    
	    for (int i = 1; i <= rowcount; i++) {
	        for (int j = 1; j <= TotalnumCol; j++) {
	        	SourceexcelLog(i, j, SourceMismatchValue);
	        }
	    }
       
	    
	    
	    
	    try {
	        FileOutputStream out = new FileOutputStream(dest);
	        myWorkBook.write(out);
	        out.close();
	    } catch (Exception e) {
	        e.printStackTrace();
	    }
	}
	
	
	public static void pushTargetDatatoExcel(int rowcount,int Totalcolumncount,String TargetMismatchValue) {
	    int TotalnumCol = Totalcolumncount; // assume 10 cols

	    for (int i = 0; i < rowcount; i++) {
	        for (int j = 0; j < TotalnumCol; j++) {
	        	TargetexcelLog(i, j, TargetMismatchValue);
	        }
	    }

	    try {
	        FileOutputStream out = new FileOutputStream(dest);
	        myWorkBook.write(out);
	        out.close();
	    } catch (Exception e) {
	        e.printStackTrace();
	    }
	}


}
