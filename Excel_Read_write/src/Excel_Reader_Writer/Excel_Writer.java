package Excel_Reader_Writer;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class Excel_Writer extends Excel_Reader{
	
	 String setCellValue(String Path_to_File,String Column_Name, int Row_Num,String Cell_value) throws RuntimeException, IOException
	{
		FileInputStream file = new FileInputStream(new File(Path_to_File));
		Workbook wb = WorkbookFactory.create(file);
		Sheet sheet = wb.getSheetAt(0);
		int col_number=getColumnNumber(Column_Name, sheet);
		
		if(col_number>=0)
		{
			
			System.out.println("Column number:"+col_number);
			Cell cell = sheet.getRow(Row_Num).createCell(col_number, CellType.STRING);
			cell.setCellValue(Cell_value);
			file.close();
			FileOutputStream outFile =new FileOutputStream(new File(Path_to_File));
			wb.write(outFile);
            outFile.close();

			
			
			return "Success";
		}
		else {
			return "Error:Target Column not found";
		}
		
	}
	
	
	
	
	
	
	
	

}
