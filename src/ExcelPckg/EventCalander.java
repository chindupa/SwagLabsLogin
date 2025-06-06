package ExcelPckg;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.By.ByLinkText;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;

public class EventCalander 
{

	public static void main(String[] args) throws IOException, InterruptedException 
	{
		String path="D:\\Selenium\\excel\\sheet\\EventCalander.xlsx";
		FileInputStream fs=new FileInputStream(path);
		XSSFWorkbook work=new XSSFWorkbook(fs);
		XSSFSheet sheet=work.getSheetAt(0);
		String Name=sheet.getRow(2).getCell(0).getStringCellValue();
		String Email=sheet.getRow(2).getCell(1).getStringCellValue();
		String Phone=sheet.getRow(2).getCell(2).getStringCellValue();
		String Password=sheet.getRow(2).getCell(3).getStringCellValue();
		String ConfirmPassword=sheet.getRow(2).getCell(4).getStringCellValue();
		
		System.setProperty("webdriver.chrome.driver", "D:\\selenium2\\Chromedriver\\chromedriver-win64\\chromedriver.exe");
		WebDriver driver=new ChromeDriver();
		driver.get("http://campus.sicsglobal.co.in/Project/EventCalendar/index.php");
		driver.manage().window().maximize();
		Thread.sleep(1000);
		JavascriptExecutor je=(JavascriptExecutor)driver;
		je.executeScript("window.scrollBy(0,800)","");
		Thread.sleep(2000);
		WebElement Register=driver.findElement(By.linkText("Register"));
		Register.click();
		Thread.sleep(1000);
		
		WebElement Nme=driver.findElement(By.name("name"));
		Nme.sendKeys(Name);
		Thread.sleep(1000);
		WebElement Eml=driver.findElement(By.name("email"));
		Eml.sendKeys(Email);
		Thread.sleep(1000);
		WebElement Phn=driver.findElement(By.name("phone"));
		Phn.sendKeys(String.valueOf(Phone));
		Thread.sleep(1000);
		WebElement Pswd=driver.findElement(By.name("password"));
		Pswd.sendKeys(Password);
		Thread.sleep(1000);
		WebElement CnfmPswd=driver.findElement(By.name("confirmpassword"));
		CnfmPswd.sendKeys(ConfirmPassword);
		Thread.sleep(2000);
		WebElement SbmtRegstr=driver.findElement(By.name("register"));
		SbmtRegstr.click();
		Thread.sleep(2000);
		WebElement Lgn=driver.findElement(By.linkText("Login"));
		Lgn.click();
		Thread.sleep(2000);
		WebElement LgnEml=driver.findElement(By.name("email"));
		LgnEml.sendKeys(Email);
		Thread.sleep(2000);
		WebElement LgnPswrd=driver.findElement(By.name("password"));
		LgnPswrd.sendKeys(Password);
		Thread.sleep(2000);
		System.out.println(Email);
		System.out.println(Password);
		WebElement SbmtLgn=driver.findElement(By.xpath("/html/body/section/div/div[2]/div[1]/form/div[3]/input"));
		SbmtLgn.click();
		Thread.sleep(2000);
		driver.quit();
	}

}
