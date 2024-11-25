package Framework.Excel;

import java.io.FileInputStream;

import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.annotations.Test;

public class dataarrayconversion{
	
	@Test
	public  Object[][] dataintoarray() throws IOException {
 		
	String path = "C:\\Users\\satishreddy.sheelam\\OneDrive - Entain Group\\Documents\\Excelautomationpractice.xlsx";
	
	FileInputStream inputStream = new FileInputStream(path);
	
	XSSFWorkbook workbook = new XSSFWorkbook(inputStream);
	
	XSSFSheet sheet = workbook.getSheet("Sheet1");	
	
	Object[][] data = new Object[sheet.getPhysicalNumberOfRows()-1][sheet.getRow(sheet.getFirstRowNum()).getLastCellNum()];
	
	for(int j=0;j<sheet.getPhysicalNumberOfRows()-1;j++) {
		
		System.out.println(sheet.getPhysicalNumberOfRows());
		
		XSSFRow row = sheet.getRow(j+1);
		
		for(int i=0; i<row.getPhysicalNumberOfCells(); i++) {
			
		data[j][i] = row.getCell(i);
		
		}
	}
	
	
	
	return data;
			
	
	}
	
	
}
