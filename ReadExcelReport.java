import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class ReadExcelReport {

	public static void main(String[] args) throws IOException {
		
		File fileread = new File("E:\\ApachePOI\\results.xlsx");
		FileInputStream inputst = new FileInputStream(fileread);
		XSSFWorkbook woor = new XSSFWorkbook(inputst);
		System.out.println(woor.getNumberOfSheets());
		XSSFSheet sheet1 = woor.getSheetAt(0);
		System.out.println(sheet1.getLastRowNum());
		XSSFRow rows = sheet1.getRow(0);
		System.out.println(rows.getFirstCellNum());
		System.out.println(rows.getLastCellNum());
		
		
		
		

	}

}
