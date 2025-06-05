package ExcelPckg;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;

public class SwaglabsLogin 
{

	public static void main(String[] args) throws IOException, InterruptedException 
	{
		String path="D:\\Selenium\\excel\\sheet\\SwaglabsLogin.xlsx";
		FileInputStream fs=new FileInputStream(path);
		XSSFWorkbook work=new XSSFWorkbook(fs);
		XSSFSheet sheet=work.getSheetAt(0);
		String Username =sheet.getRow(2).getCell(0).getStringCellValue();
		String Password=sheet.getRow(2).getCell(1).getStringCellValue();
		
		System.setProperty("webdriver.chrome.driver", "D:\\selenium2\\Chromedriver\\chromedriver-win64\\chromedriver.exe");
		WebDriver driver=new ChromeDriver();
		driver.get("https://www.saucedemo.com/v1/");
		driver.manage().window().maximize();
		WebElement UsrNme=driver.findElement(By.id("user-name"));
		UsrNme.sendKeys(Username);
		Thread.sleep(1000);
		WebElement Psswrd=driver.findElement(By.id("password"));
		Psswrd.sendKeys(Password);
		Thread.sleep(1000);
		WebElement Lgn=driver.findElement(By.id("login-button"));
		Lgn.click();
		Thread.sleep(1000);
		String ExpectedTitle="Swag Labs";
		String ActualTitle=driver.getTitle();
		if(ExpectedTitle.equalsIgnoreCase(ActualTitle))
		{
			System.out.println("Login test successfully");
		}
		else
		{
			System.out.println("Login test Failed");
		}
		Thread.sleep(2000);
		driver.quit();
	}

}
