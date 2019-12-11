package org.cts.test;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class GitHubConflict {
	public static void main(String[] args) throws IOException

	{
	File f= new File("C:\\Users\\PMV\\Desktop\\Master Clone\\AirCanadaSez\\TestData\\Excel_Read.xlsx");
	FileInputStream a=new FileInputStream(f);
	Workbook w= new XSSFWorkbook(a);
	
	Sheet s=w.getSheet("Sheet1");
	Row r=s.getRow(3);
	Cell c=r.getCell(3);
	
	System.out.println(c);
	
	int nr=s.getPhysicalNumberOfRows();
	
	System.out.println("Number of Row: "+nr);
	
	int nc=r.getPhysicalNumberOfCells();
	
	System.out.println("Number of Columns: "+nc);
	
	for(int i=0;i<s.getPhysicalNumberOfRows();i++)
		
	{
		Row r1=s.getRow(i);
		
		for(int j=0;j<r1.getPhysicalNumberOfCells();j++)
			
		{
		
			Cell c1=r1.getCell(j);
			
			System.out.println(c1);
		
	
	}	
	
	}
	//Code changes made by Teammember1
	System.out.println("Code added for Excel Write");
}

}
