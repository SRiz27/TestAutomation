package Excel_Reader_Writer;

import java.io.FileInputStream;
import java.io.IOException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class Excel_Reader {
	
	String getCellValue(String Path_to_File,String Column_Name, int Row_Num) throws RuntimeException, IOException
	{
		FileInputStream file = new FileInputStream(Path_to_File); 
		Workbook wb = WorkbookFactory.create(file);
		Sheet sheet = wb.getSheetAt(0);
		int col_number=getColumnNumber(Column_Name, sheet);
				if(col_number>=0)
				{
					Row targetRow = sheet.getRow(Row_Num);
					Cell targetCell=targetRow.getCell(col_number)	;
//					CellType type =targetCell.getCellType();
					String value = targetCell.getStringCellValue();
					wb.close();
					file.close();
					return value;
					
				}
				else {
					return "Sorry,Target Column not found";
				}
		
	}
	
	int getColumnNumber(String ColumnName,Sheet Sheet)
	{

		Row firstRow = Sheet.getRow(0);
		
		for (int cellIndex = 0; cellIndex < firstRow.getLastCellNum(); cellIndex++)
		{
		    Cell cell = firstRow.getCell(cellIndex);
		    String cellValue = cell.getStringCellValue(); 	    
		    if(cellValue.equalsIgnoreCase(ColumnName))
		    {
		    	System.out.println("Cell " + cellIndex + ": " + cellValue);
		    	return cellIndex;
		    }
		    
			
		}
		
		return -1;
		
		
		
		
		
		
	}
	
	
	
	
	
	
	

}
