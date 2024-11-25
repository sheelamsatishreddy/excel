package Framework.Excel;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.sl.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.annotations.Test;

public class Readdatafromexcel {

	
	/*We Have 4 interfaces in java to interact with excel files
	
		- Workbook Interface - this interface methods are again defined by two classes. One class for xls extension workbooks and another for xlsx extension workbooks. Depending on the type of workbook we need to use the class
		- Sheet Interface - this interface methods are again defined by two classes. One class for xls extension workbooks and another for xlsx extension workbooks. Depending on the type of workbook we need to use the class
		- Row Interface - this interface methods are again defined by two classes. One class for xls extension workbooks and another for xlsx extension workbooks. Depending on the type of workbook we need to use the class
		- Cell Interface - this interface methods are again defined by two classes. One class for xls extension workbooks and another for xlsx extension workbooks. Depending on the type of workbook we need to use the class
	
		- workbook.iterator()-Helps to iterate over all the sheets
		- sheet.iterator()-Helps to iterate over all rows
		- row.iterator()-Helps to iterate over all cells
	
	*/
	
	@Test
	public void getExcelData() throws IOException {
		
		
		String path = "C:\\Users\\satishreddy.sheelam\\OneDrive - Entain Group\\Documents\\Excelautomationpractice.xlsx";
		
		FileInputStream inputStream = new FileInputStream(path);;
		
		XSSFWorkbook workbook = new XSSFWorkbook(inputStream);
		
		XSSFSheet sheet = workbook.getSheet("Sheet1");	
		
		for(int j=0;j<sheet.getPhysicalNumberOfRows();j++) {
			
			XSSFRow row = sheet.getRow(j);
			
			for(int i=0; i<row.getPhysicalNumberOfCells(); i++) {
				
				System.out.println(row.getCell(i));
			
			}
		}
		
	}
		
		
		
		
		
		
		
		
		//cell iterator reads only the values defined cells will not read the exmpty cells
		

}

