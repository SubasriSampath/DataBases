package com.simple;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.math.BigDecimal;
import java.text.SimpleDateFormat;
import java.util.Date;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class FamilyDatabase {
public static void main(String[] args) throws IOException {
	File file = new File("C:\\Users\\aaa\\eclipse-workspace\\DataBases\\Excel Sheet\\Data.xlsx");
	FileInputStream stream = new FileInputStream(file);
	Workbook workbook = new XSSFWorkbook(stream);
	Sheet sheet = workbook.getSheet("Data");
	//Row row = sheet.getRow(3);
	//Cell cell = row.getCell(3);
	//System.out.println(cell);
	for(int i=0; i<sheet.getPhysicalNumberOfRows(); i++)
	{
		Row row =sheet.getRow(i);
		for(int j=0; j<row.getPhysicalNumberOfCells(); j++)
		{
			Cell cell = row.getCell(j);
			CellType cellType = cell.getCellType();
			
			switch(cellType) {
			case STRING:
				String name = cell.getStringCellValue();
				System.out.println(name);
				break;
				
			case NUMERIC:
			if(	DateUtil.isCellDateFormatted(cell))
			{
				Date dateCellValue = cell.getDateCellValue();
				SimpleDateFormat dtformat = new SimpleDateFormat("dd/MMM/yyyy");
				String format = dtformat.format(dateCellValue);
				System.out.println(format);
			}
			else
			{
				double d = cell.getNumericCellValue();
				BigDecimal b = BigDecimal.valueOf(d);
				String num = b.toString();
				System.out.println(num);
			}
			break;
			default:
				break;
		}
	}
	}
	workbook.close();
}
}
