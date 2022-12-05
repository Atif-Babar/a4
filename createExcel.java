package Excel;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class createExcel {

	public static void main(String[] args) throws Exception {
		File excelFile = new File("./src/test/resources/Test.xlsx");
		FileInputStream fis = new FileInputStream(excelFile);
		XSSFWorkbook workbook = new XSSFWorkbook(fis);
		XSSFSheet sheet = workbook.getSheetAt(0);
		
		int noOfRows = sheet.getPhysicalNumberOfRows();
		System.out.println("Number of Rows: " + noOfRows);            // row count	
		
		sheet.getRow(0).createCell(2).setCellValue("Status");
		sheet.getRow(1).createCell(2).setCellValue("Valid");
		sheet.getRow(2).createCell(2).setCellValue("Invalid");
		sheet.getRow(3).createCell(2).setCellValue("Invalid");
		
		System.out.println("we have write the data on excel");
		
		FileOutputStream  fout=new FileOutputStream(excelFile);
		workbook.write(fout);
		workbook.close();
		

	}

}
