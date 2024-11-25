package Framework.Excel;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import javax.swing.text.html.HTMLDocument.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.annotations.Test;

public class ReadingdatafromExcelusingIterator {

	@Test
	public void readDataIteratorexcel() throws IOException {
		
		String path = "C:\\Users\\satishreddy.sheelam\\OneDrive - Entain Group\\Documents\\Excelautomationpractice.xlsx";
		
		FileInputStream inputStream = new FileInputStream(path);
		
		XSSFWorkbook workbook = new XSSFWorkbook(inputStream);
		
		java.util.Iterator<Sheet> sheets = workbook.iterator();
				
		while(sheets.hasNext()) {
			
		XSSFSheet sheet = (XSSFSheet) sheets.next();
		
		if(sheet.getSheetName().equalsIgnoreCase("Sheet1")) {
		
		java.util.Iterator<Row> rows =	sheet.iterator();
		
		Object[][] data;
		
		while(rows.hasNext()) {
			
			XSSFRow row = (XSSFRow) rows.next();
			
			java.util.Iterator<Cell> cells = row.cellIterator();
			
			while(cells.hasNext()) {
				
				Cell cell = cells.next();
			
				switch(cell.getCellType()){
					
				case STRING: System.out.print(cell.getStringCellValue());break;
				case NUMERIC: System.out.print(cell.getNumericCellValue());break;
					
				}
				
				System.out.print("  ||  ");	
			}
			
			System.out.println();
		
		}
		
		}
			
		}
			
	}
	
}
