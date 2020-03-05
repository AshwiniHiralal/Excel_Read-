package Feb15;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

public class WriteToExistingFile {

	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub

		
		
      File f = new File("C:\\Users\\INDIA\\Documents\\ExcelReadData.xls");
		
		FileInputStream input =new FileInputStream(f);
		
		HSSFWorkbook workbook = new HSSFWorkbook(input);
		
		HSSFSheet sheet = workbook.getSheet("Sheet1");
		
		
		HSSFRow row=sheet.getRow(0);
		int maxCell= row.getLastCellNum();
		HSSFCell cell= row.createCell(maxCell);
		cell.setCellValue("Dhule");
		
		
		
		int maxRow=sheet.getLastRowNum();
		HSSFRow row1= sheet.createRow(maxRow+1);
		
		int maxCell1= row1.getLastCellNum();
		
		HSSFCell cell1=row1.createCell(maxCell1+1);
		
		cell1.setCellValue("OOty");
		
		
		FileOutputStream output = new FileOutputStream(f);
		
		workbook.write(output);
		output.close();
		
	
		
	}

}
