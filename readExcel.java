package Excel;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;

import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class readExcel {

	public static void main(String[] args) throws Exception {
		File excelFile = new File("./src/test/resources/Test.xlsx");
		FileInputStream fis = new FileInputStream(excelFile);
		XSSFWorkbook workbook = new XSSFWorkbook(fis);
		XSSFSheet sheet = workbook.getSheet("Sheet1");
		
		int noOfRows = sheet.getPhysicalNumberOfRows(); 
		int noOfColumns = sheet.getRow(0).getLastCellNum();
		System.out.println("Number of Rows: " + noOfRows);            // row count		
		
		
		String[][] data = new String[noOfRows-1][noOfColumns];
		for (int i = 1; i < noOfRows; i++) {
			for (int j = 0; j < noOfColumns; j++) {
				DataFormatter df = new DataFormatter();
				System.out.println(sheet.getRow(i).getCell(j).getStringCellValue()) ;
			}
			System.out.println();
		}
		workbook.close();
		fis.close();

	}

}
