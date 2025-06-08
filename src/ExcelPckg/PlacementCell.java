package ExcelPckg;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.Select;

public class PlacementCell {

	public static void main(String[] args) throws IOException, InterruptedException 
	{
		String path="D:\\Selenium\\excel\\sheet\\PlacementCell.xlsx";
		FileInputStream fs= new FileInputStream(path);
		XSSFWorkbook work= new XSSFWorkbook(fs);
		XSSFSheet sheet= work.getSheetAt(0);
		XSSFRow row=sheet.getRow(2);
		System.setProperty("webdriver.chrome.driver", "D:\\selenium2\\Chromedriver\\chromedriver-win64\\chromedriver.exe");
		WebDriver driver=new ChromeDriver();
		driver.get("http://campus.sicsglobal.co.in/Project/placementcell/index.php");
		driver.manage().window().maximize();
		Thread.sleep(2000);
		
		  WebElement RgstrNw=driver.findElement(By.linkText("Register Now"));
		  RgstrNw.click();
		  Thread.sleep(2000);
		  String AdmNumber=row.getCell(0).getStringCellValue();
		  WebElement AdmnNo=driver.findElement(By.name("admission_no"));
		  AdmnNo.sendKeys(AdmNumber);
		  Thread.sleep(1000);
		  WebElement Dept=driver.findElement(By.name("department"));
		  Select SelDept=new Select(Dept);
		  SelDept.selectByIndex(1);
		  Thread.sleep(1000);
		  String Batchname=row.getCell(1).getStringCellValue();
		  WebElement Btch=driver.findElement(By.name("batch"));
		  Btch.sendKeys(Batchname);
		  Thread.sleep(1000);
		  String CompltdYear=row.getCell(2).getStringCellValue();
		  WebElement CrsCpltnYr=driver.findElement(By.name("completion_year"));
		  CrsCpltnYr.sendKeys(CompltdYear);
		  Thread.sleep(1000);
		  String Name=row.getCell(3).getStringCellValue();
		  WebElement Nme=driver.findElement(By.name("name"));
		  Nme.sendKeys(Name);
		  Thread.sleep(1000);
		  String DoBrth=row.getCell(4).getStringCellValue();
		  WebElement Dob= driver.findElement(By.xpath("/html/body/div/section[1]/div[2]/div/div[2]/form/div[6]/input")); 
		  Dob.sendKeys(DoBrth);
		  String Phone=row.getCell(5).getStringCellValue();
		  WebElement Phn=driver.findElement(By.name("phone"));
		  Phn.sendKeys(Phone);
		  Thread.sleep(1000);
		  String Gender=row.getCell(6).getStringCellValue();
		  if(Gender.equalsIgnoreCase("Male"))
		  {
		  WebElement Mle=driver.findElement(By.id("male"));
		  Mle.click();
		  }
		  else
		  {
			  WebElement Fmle=driver.findElement(By.id("female"));
			  Fmle.click();
		  }
		  Thread.sleep(1000);
		  String Email=row.getCell(7).getStringCellValue();
		  WebElement Eml=driver.findElement(By.name("email"));
		  Eml.sendKeys(Email);
		  Thread.sleep(1000);
		  String Address=row.getCell(8).getStringCellValue();
		  WebElement Addrs=driver.findElement(By.name("address"));
		  Addrs.sendKeys(Address);
		  Thread.sleep(1000);
		  String Password=row.getCell(9).getStringCellValue();
		  WebElement Pswrd=driver.findElement(By.name("password"));
		  Pswrd.sendKeys(Password);
		  Thread.sleep(1000);
		  JavascriptExecutor je=(JavascriptExecutor)driver;
		  je.executeScript("window.scrollBy(0,500)", "");
		  String Image=row.getCell(10).getStringCellValue();
		  System.out.println(Image);
		  WebElement Phto=driver.findElement(By.xpath("/html/body/div/section[1]/div[2]/div/div[2]/form/div[12]/input"));
		  Phto.sendKeys(Image);
		  Thread.sleep(1000);
		  WebElement Rsme=driver.findElement(By.xpath("/html/body/div/section[1]/div[2]/div/div[2]/form/div[13]/input"));
		  Rsme.sendKeys(Image);
		  Thread.sleep(2000);
		  WebElement Rgstr=driver.findElement(By.name("add"));
		  Rgstr.click();
		  Thread.sleep(1000);
		  driver.quit();

	}

}
