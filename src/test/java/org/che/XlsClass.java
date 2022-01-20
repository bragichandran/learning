package org.che;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Date;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;



public class XlsClass {
public static void main(String[] args) throws IOException {
	File file = new File("C:\\Users\\DELL\\eclipse-workspace\\Lxs\\lib\\mavenEx.xlsx");
	FileInputStream stream	= new FileInputStream(file);
	Workbook workbook = new XSSFWorkbook(stream);
	Sheet sheet = workbook.getSheet("Data");
	for (int i = 0; i < sheet.getPhysicalNumberOfRows()-1; i++) {
		Row row = sheet.getRow(i);
		for (int j = 0; j <  row.getPhysicalNumberOfCells()-1; j++) {
			Cell cell = row.getCell(j);
			

			int type = cell.getCellType();
			if (type==1) {
				String cellValue = cell.getStringCellValue();
				System.out.println(cellValue);
			}
			
			if(type==0) {
				
				if (DateUtil.isCellDateFormatted(cell)) {
					Date dateCellValue = cell.getDateCellValue();
					SimpleDateFormat dateFormat = new SimpleDateFormat("dd-MMM-yy");
					String data1 = dateFormat.format(dateCellValue);
					System.out.println(data1);
				}
				double numericCellValue = cell.getNumericCellValue();
				long l = (long)numericCellValue;
				
				String valueOf = String.valueOf(l);
				System.out.println(valueOf);
		
		
		
		
		}
	
	}
	
	
	
	
	
}




}
}