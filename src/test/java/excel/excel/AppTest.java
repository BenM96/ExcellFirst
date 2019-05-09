package excel.excel;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Row.MissingCellPolicy;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;

/**
 * Unit test for simple App.
 */
public class AppTest{
	
	@Test
	public void test1() throws IOException {
		FileInputStream file= new FileInputStream(Constant.TEST);
		XSSFWorkbook workbook = new XSSFWorkbook(file);
		XSSFSheet sheet = workbook.getSheetAt(0);
		
		//Reading
		for(int row=0;row<sheet.getPhysicalNumberOfRows();row++) {
			for(int col=0; col<sheet.getRow(row).getPhysicalNumberOfCells();col++) {		
				XSSFCell cell=sheet.getRow(row).getCell(col);
				String cellVal=cell.toString();
				System.out.println(cellVal);
			}
		}
		
		XSSFRow row=sheet.createRow(3);
		XSSFCell cell=row.getCell(0,MissingCellPolicy.RETURN_BLANK_AS_NULL);
		cell=row.createCell(0);
		cell.setCellValue("hello");
		
		
		FileOutputStream fileOut = new FileOutputStream(Constant.TEST);
		workbook.write(fileOut);
		fileOut.flush();
		fileOut.close();
		file.close();
		
		
		
	}
    
}
