package ExcelPckg;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.formula.EvaluationWorkbook;
import org.apache.poi.ss.formula.udf.UDFFinder;
import org.apache.poi.ss.usermodel.CellReferenceType;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Name;
import org.apache.poi.ss.usermodel.PictureData;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.SheetVisibility;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.Select;
import org.apache.poi.ss.usermodel.Row.MissingCellPolicy;

public class TripPlanner {

	public static void main(String[] args) throws IOException, InterruptedException 
	{
		String path="D:\\Selenium\\excel\\sheet\\TripPlanner.xlsx";
		FileInputStream fs= new FileInputStream(path);
		XSSFWorkbook work= new XSSFWorkbook(fs);
		XSSFSheet sheet= work.getSheetAt(0);
		String Name=sheet.getRow(2).getCell(0).getStringCellValue();
		String Age=sheet.getRow(2).getCell(1).getStringCellValue();
		String Phone=sheet.getRow(2).getCell(2).getStringCellValue();
		String Email=sheet.getRow(2).getCell(3).getStringCellValue();
		String Password=sheet.getRow(2).getCell(4).getStringCellValue();
			
		System.setProperty("webdriver.chrome.driver", "D:\\selenium2\\Chromedriver\\chromedriver-win64\\chromedriver.exe");
		WebDriver driver=new ChromeDriver();
		driver.get("http://campus.sicsglobal.co.in/Project/tripplanner/index.php");
		driver.manage().window().maximize();
		Thread.sleep(2000);
		WebElement Reg=driver.findElement(By.xpath("/html/body/div/div/div/div[2]/nav/ul/li[3]/a"));
		Reg.click();
		Thread.sleep(2000);
		WebElement State=driver.findElement(By.name("state"));
		Select SelState=new Select(State);
		SelState.selectByIndex(1);
		Thread.sleep(1000);
		WebElement District=driver.findElement(By.name("district"));
		Select SelDist=new Select(District);
		SelDist.selectByIndex(1);
		Thread.sleep(1000);
		WebElement Nme=driver.findElement(By.name("name"));
		Nme.sendKeys(Name);
		Thread.sleep(1000);
		WebElement Ag=driver.findElement(By.name("age"));
		Ag.sendKeys(Age);
		Thread.sleep(1000);
		WebElement Phn=driver.findElement(By.name("phone"));
		Phn.sendKeys(Phone);
		Thread.sleep(1000);
		WebElement Eml=driver.findElement(By.name("email"));
		Eml.sendKeys(Email);
		Thread.sleep(1000);
		WebElement Pswrd=driver.findElement(By.name("password"));
		Pswrd.sendKeys(Password);
		Thread.sleep(1000);
		
		WebElement SubtRegstr=driver.findElement(By.xpath("//*[@id=\"form1\"]/div[2]/div[8]/input[1]"));
		SubtRegstr.click();
		Thread.sleep(2000);
		WebElement LogUser=driver.findElement(By.linkText("Skip to Login"));
		LogUser.click();
		Thread.sleep(1000);
		WebElement Usrpswrd=driver.findElement(By.xpath("//*[@id=\"contact_name\"]"));
		Usrpswrd.sendKeys(Password);
		Thread.sleep(1000);
		WebElement Login=driver.findElement(By.xpath("//*[@id=\"form1\"]/div[3]/input[1]"));
		Login.click();
		Thread.sleep(2000);
		
	}
			
	

}
