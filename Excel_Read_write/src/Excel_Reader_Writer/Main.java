package Excel_Reader_Writer;

import java.io.IOException;

public class Main {

	public static void main(String[] args) throws RuntimeException, IOException {
		// TODO Auto-generated method stub
		String Excel_Path="C:\\Users\\ShadabRizvi\\Desktop\\Java_practice.xls";
		
		
//		
//		Excel_Reader er = new Excel_Reader();
//		
//		System.out.println("Cell value found:"+er.getCellValue(Excel_Path,"Output_Data",3));
		Excel_Writer	ew = new Excel_Writer();
		ew.setCellValue(Excel_Path,"Result", 1,"Pass");
		
		
		

	}

}
