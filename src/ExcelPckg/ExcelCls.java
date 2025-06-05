package ExcelPckg;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelCls {

	public static void main(String[] args) throws IOException 
	{
		String path="D:\\Selenium\\excel\\sheet\\Fruitsnew.xlsx";
		FileInputStream fs= new FileInputStream(path);
		XSSFWorkbook work=new XSSFWorkbook(fs);
		XSSFSheet sheet=work.getSheetAt(0);
		XSSFRow row= sheet.getRow(0);
		XSSFCell password=row.getCell(0);
		for(int i=2;i<=6;i++)
		{
        System.out.println(sheet.getRow(i).getCell(0));
		}

	}

}
