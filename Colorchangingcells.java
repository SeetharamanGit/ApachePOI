import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Hyperlink;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Colorchangingcells {

	public static void main(String[] args) throws IOException {

		File fileread = new File("E:\\ApachePOI\\results.xlsx");
		FileInputStream inputst = new FileInputStream(fileread);

		XSSFWorkbook woor = new XSSFWorkbook(inputst);
		CreationHelper helper =  woor.getCreationHelper();
		Hyperlink link = helper.createHyperlink(Hyperlink.LINK_URL);
		link.setAddress("https://www.google.co.in");
		XSSFSheet sheet1 = woor.getSheetAt(0);
	
		int rownum = sheet1.getLastRowNum();

		for (Row row1 : sheet1) {

			System.out.println(row1.getCell(2));
			String Teststatus = row1.getCell(2).getStringCellValue();
			//System.out.println(Teststatus);
			

			if (Teststatus.equals("Pass")) {
				System.out.println(Teststatus);
				CellStyle style = woor.createCellStyle();
				style.setFillPattern(CellStyle.DIAMONDS);
				style.setFillBackgroundColor(IndexedColors.GREEN.getIndex());
				
				//style.setFillForegroundColor(IndexedColors.GREEN.getIndex());
				row1.getCell(2).setCellStyle(style);
				row1.getCell(2).setHyperlink(link);

			}

			if (Teststatus.equals("Fail")) {
				System.out.println(Teststatus);
				CellStyle style = woor.createCellStyle();
				style.setFillPattern(CellStyle.BIG_SPOTS);
				style.setFillBackgroundColor(IndexedColors.RED.getIndex());
				//style.setFillPattern(CellStyle.ALT_BARS);
				//style.setFillForegroundColor(IndexedColors.RED.getIndex());
				row1.getCell(2).setCellStyle(style);

			}
		}
		
		FileOutputStream fileOut = new FileOutputStream(fileread);
	    woor.write(fileOut);
	    fileOut.flush();
	    fileOut.close();
		
	}

}
