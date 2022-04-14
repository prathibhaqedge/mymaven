package Driverfactory;

import java.util.concurrent.TimeUnit;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.testng.Reporter;
import org.testng.annotations.AfterTest;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.Test;

import com.relevantcodes.extentreports.ExtentReports;
import com.relevantcodes.extentreports.ExtentTest;
import com.relevantcodes.extentreports.LogStatus;

import Utilites.ExcelUtilities;

public class AppTest {
	WebDriver driver;
	String inputpath="E:\\workspace\\Orange_Maven\\TestInput\\loginData.xlsx";
	String outputpath="E:\\workspace\\Orange_Maven\\TestOutput\\resultogloginorange.xlsx";
	ExtentReports reports;
	ExtentTest test;
	@BeforeTest
	public void setup()
	{
		reports = new ExtentReports("./ExtentReport/drivermavenresult.html");
		driver = new ChromeDriver();
		driver.manage().window().maximize();
		
	}
	@Test
	public void validatelogin()throws Throwable 
	{
		ExcelUtilities xl = new ExcelUtilities(inputpath);
		int rc =xl.rowcount("Login");
		Reporter.log("no of rows are :::::::"+rc,true);
		for(int i=1;i<=rc;i++)
		{
			test= reports.startTest("validate login");
		driver.get("http://www.orangehrm.qedgetech.com/");
		driver.manage().deleteAllCookies();
		driver.manage().timeouts().implicitlyWait(30,TimeUnit.SECONDS);
		String username = xl.getcelldata("Login", i, 0);
		String password=xl.getcelldata("Login", i, 1);
		driver.findElement(By.name("txtUsername")).sendKeys(username);
		driver.findElement(By.name("txtPassword")).sendKeys(password);
		driver.findElement(By.name("Submit")).click();
		String expected = "dashboard";
		String actual=driver.getCurrentUrl();
		if(actual.contains(expected))
		{
			Reporter.log("login sucess ::::::"+expected+"    "+actual);
			xl.setcelldata("Login", i, 2, "login sucess", outputpath);
			xl.setcelldata("Login", i, 3, "pass", outputpath );
			test.log(LogStatus.PASS, "login sucess::::::"+expected);
		}
		else
		{
			Reporter.log("login fail ::::::"+expected+"    "+actual);
			xl.setcelldata("Login", i, 2, "login fail", outputpath);
			xl.setcelldata("Login", i, 3, " fail", outputpath );
			test.log(LogStatus.FAIL, "login fail::::::"+expected);
		}
		reports.endTest(test);
		reports.flush();
		
		
		
	}
	}
	@AfterTest
	public void teardown()
	{
		driver.close();
	}
	
	

}
