package Feb16;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

public class ValidateAndReadData {

	public static void main(String[] args) throws IOException{
		// TODO Auto-generated method stub
		
	File f = new File("C:\\Users\\INDIA\\Documents\\ExcelAllType.xls");
		
		FileInputStream input =new FileInputStream(f);
		
        HSSFWorkbook workbook = new HSSFWorkbook(input);
        
		
		HSSFSheet sheet = workbook.getSheet("Sheet1");
	
		HSSFRow row=sheet.getRow(3);
		HSSFCell cell= row.getCell(2);//we count from 0 & getLastCell method count from 1
		
		
		if(cell==null)
		{
			
			System.out.println("Cell is NULL");
		}
		else if(cell.getCellType()== HSSFCell.CELL_TYPE_BLANK)
		{
			System.out.println("Cell is the blank");
		}
		else if(cell.getCellType()== HSSFCell.CELL_TYPE_BOOLEAN)
		{
			System.out.println(cell.getBooleanCellValue());
		}
		else if(cell.getCellType()== HSSFCell.CELL_TYPE_NUMERIC)
		{
			System.out.println(cell.getNumericCellValue());
		}
		else if(cell.getCellType()== HSSFCell.CELL_TYPE_STRING)
		{
			System.out.println(cell.getStringCellValue());
		}

	}
	

}
