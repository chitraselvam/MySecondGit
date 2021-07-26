package org.testdata;

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

public class TestFile {
	public static void main(String[] args) throws IOException {
		File newFile= new File("C:\\\\Users\\\\Dell\\\\Documents\\\\Testdata_File.xlsx");
		FileInputStream filestream=new FileInputStream(newFile);
		Workbook book= new XSSFWorkbook(filestream);
		Sheet sheet = book.getSheet("Sheet1");
		for(int i=0;i<sheet.getPhysicalNumberOfRows();i++) {
			Row row = sheet.getRow(i);
			for(int j=0;j<row.getPhysicalNumberOfCells();j++)
			{
				Cell cell = row.getCell(j);
				int cellType = cell.getCellType();
				if(cellType==1)
				{
				String stringCellValue = cell.getStringCellValue();
				System.out.println(stringCellValue);
				}
				else if(DateUtil.isCellDateFormatted(cell))
				{
					Date dateCellValue = cell.getDateCellValue();
					SimpleDateFormat format=new SimpleDateFormat("MMM-dd-YYYY");
					String format2 = format.format(dateCellValue);
					System.out.println(format2);
					
					
				}
				
				else
				{
					double numericCellValue = cell.getNumericCellValue();
					long l=(long)numericCellValue;
					System.out.println(l);
				}
				}
			}
		}
		
		
		
	

}
