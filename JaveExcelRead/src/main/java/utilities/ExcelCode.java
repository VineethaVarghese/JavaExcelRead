package utilities;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;



public class ExcelCode {
	
	static FileInputStream f;
	static XSSFWorkbook w;
	static XSSFSheet sh;
	
public static String getStringData(int a , int b) throws IOException
//throws,there maybe chances of an exception to occur 
{
	f=new FileInputStream("C:\\Users\\shon4\\eclipse-workspace\\JaveExcelRead\\src\\main\\resources\\ExcelTest.xlsx");
	
	w=new XSSFWorkbook(f);//Inbuilt class ,excel workbook
	sh=w.getSheet("Sheet1");//getSheet is a method, to select which excel sheet frm workbook
	Row r=sh.getRow(a);//Row,cell bth are interface. getRow ,getCell they are method
	Cell c=r.getCell(b);//to get the cell value from the row
	//Here we have parameterized row & cell values
	//public static String ,,its a return type
	return c.getStringCellValue();

}
public static String getIntegerData(int a , int b)  throws IOException
{
	f=new FileInputStream("C:\\Users\\shon4\\eclipse-workspace\\JaveExcelRead\\src\\main\\resources\\ExcelTest.xlsx");
	w=new XSSFWorkbook(f);
	sh=w.getSheet("Sheet1");
	Row r=sh.getRow(a);
	Cell c=r.getCell(b);
	int x= (int) c.getNumericCellValue();
//we want the value to be integer,so if any double or float comes it converts them to integer
	return String.valueOf(x);//valueOf() , method to convert integer value to String
}
	

}
