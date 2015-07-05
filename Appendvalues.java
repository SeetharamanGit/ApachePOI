import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class Appendvalues {

	public static void main(String[] args) throws IOException {


		File fileread = new File("E:\\ApachePOI\\results.xlsx");
		FileInputStream inputst = new FileInputStream(fileread);
		XSSFWorkbook woor = new XSSFWorkbook(inputst);
		
		XSSFSheet sheet1 = woor.getSheetAt(0);
		int rows = sheet1.getLastRowNum();
		
		FileOutputStream fo = new FileOutputStream(fileread);
		
		XSSFRow rowsappende = sheet1.createRow(rows+1);
		rowsappende.createCell(0).setCellValue("3");
		rowsappende.createCell(1).setCellValue("Login2");
		rowsappende.createCell(2).setCellValue("Fail");
		
		woor.write(fo);
		
		System.out.println("For Git Update");
		System.out.println("For Git Branch");
		
		
		
		

	}

}
