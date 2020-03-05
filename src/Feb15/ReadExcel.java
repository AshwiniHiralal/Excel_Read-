package Feb15;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;


public class ReadExcel {

	public static void main(String[] args) throws IOException{
		// TODO Auto-generated method stub

		
		
		File f = new File("C:\\Users\\INDIA\\Documents\\ExcelReadData.xls");
		
		FileInputStream input =new FileInputStream(f);
		
		HSSFWorkbook workbook = new HSSFWorkbook(input);
		
		HSSFSheet sheet = workbook.getSheet("Sheet1");
		
		System.out.println(sheet.getLastRowNum());
		HashMap<Integer, ArrayList<String>> map = new HashMap<>();
		
		
		for(int i=0; i<= sheet.getLastRowNum(); i++)
		{
			HSSFRow row= sheet.getRow(i);
			
			ArrayList<String> list= new ArrayList<>();
			//need to declare whenewer iteration changes
			for(int j=0; j<row.getLastCellNum();j++)
			{
				HSSFCell cell = row.getCell(j);
		   	  System.out.print(cell.getStringCellValue()+ " ");
			  list.add(cell.getStringCellValue());
			  	
				
			}
			System.out.println();
			map.put(i, list);
			
		}
		System.out.println(map);
		
	}

}
