package Feb15;

import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

public class CreateNewFile {

	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub

		
		
		
HSSFWorkbook workbook = new HSSFWorkbook();
		
		HSSFSheet sheet = workbook.createSheet("NewSheet");
		
		
		
		int maxRow = sheet.getLastRowNum();
		HSSFRow row=sheet.createRow(maxRow);
		int maxCell= row.getLastCellNum();
		HSSFCell cell= row.createCell(maxCell+1);
		cell.setCellValue("Dhule");
		
		
		
		int maxCell2= row.getLastCellNum();
		HSSFCell cell2= row.createCell(maxCell2);
		cell2.setCellValue("OOTY");
		
		
		
FileOutputStream output = new FileOutputStream("C:\\Users\\INDIA\\Documents\\NewFile.xls");
		
		workbook.write(output);
		output.close();
		
	}

}
