import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class createexcelrreport {

	public static void main(String[] args) throws IOException {
		
		File excelwrite = new File("E:\\ApachePOI\\results.xlsx");
		FileOutputStream fo = new FileOutputStream(excelwrite);
		XSSFWorkbook wb = new XSSFWorkbook();
		XSSFSheet sheetw = wb.createSheet("TestResults");
		XSSFRow rows = sheetw.createRow(0);
		rows.createCell(0).setCellValue("S.No");
		rows.createCell(1).setCellValue("Test Case Name");
		rows.createCell(2).setCellValue("Status");
		
		wb.write(fo);
		
		
		
		
	}

}
