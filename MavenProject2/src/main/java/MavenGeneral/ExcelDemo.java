package MavenGeneral;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
 
public class ExcelDemo {
	static FileInputStream fc;
	static XSSFWorkbook wb;
	static XSSFSheet sh;
	
	
	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub
  // int data=getIntData(0,1);
   //System.out.println(data);
		int data1=getIntData(1,0);
		System.out.println(data1);
	}
	
	public static String getStringData(int a, int b) throws IOException {
	fc=new FileInputStream("D:\\Test.xlsx");
	wb=new XSSFWorkbook(fc);
	sh=wb.getSheet("Sheet1");
	Row r=sh.getRow(a);
	Cell c=r.getCell(b);
	return c.getStringCellValue();
	
	}
	
	public static int getIntData(int a, int b) throws IOException {
		fc=new FileInputStream("D:\\Test.xlsx");
		wb=new XSSFWorkbook(fc);
		sh=wb.getSheet("Sheet1");
		Row r=sh.getRow(a);
		Cell c=r.getCell(b);
		return (int)c.getNumericCellValue();
		
		}
	}
