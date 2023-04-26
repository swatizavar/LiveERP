package commonFunctions;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Properties;

import org.openqa.selenium.By;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;

public class FunctionLibrary {
public static WebDriver driver;
public static Properties conpro;
public static String Expected="";
public static String Actual="";

//Method for launch browser
public static WebDriver startBrowser() throws Throwable   
{
	conpro=new Properties();
	conpro.load(new FileInputStream("./PropertyFile/Environment.properties"));
	if(conpro.getProperty("Browser").equalsIgnoreCase("chrome"))
	{
		driver=new ChromeDriver();
		driver.manage().window().maximize();
		driver.manage().deleteAllCookies();
	}
	else if(conpro.getProperty("Browser").equalsIgnoreCase("firefox"))
	{
		driver=new FirefoxDriver();
		driver.manage().deleteAllCookies();
	}
	else
	{
		System.out.println("Browser value not matching");
	}
	return driver;
}

//Method to launch url
public static void openUrl(WebDriver driver)
{
	driver.get(conpro.getProperty("Url"));
}

//method for waitForElement
public static void waitForElement(WebDriver driver,String LocatorType, String LocatorValue, String waitTime)
{
	WebDriverWait myWait= new WebDriverWait(driver, Integer.parseInt(waitTime));
	if(LocatorType.equalsIgnoreCase("name"))
	{
		myWait.until(ExpectedConditions.visibilityOfElementLocated(By.name(LocatorValue)));
	}
	else if(LocatorType.equalsIgnoreCase("xpath"))
	{
		myWait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(LocatorValue)));
	}
	else if(LocatorType.equalsIgnoreCase("id"))
	{
		myWait.until(ExpectedConditions.visibilityOfElementLocated(By.id(LocatorValue)));
	}
}

//method for typeAction
public static void typeAction(WebDriver driver,String LocatorType, String LocatorValue,String TestData)
{
	if(LocatorType.equalsIgnoreCase("id"))
	{
		driver.findElement(By.id(LocatorValue)).clear();
		driver.findElement(By.id(LocatorValue)).sendKeys(TestData);
	}
	else if(LocatorType.equalsIgnoreCase("xpath"))
	{
		driver.findElement(By.xpath(LocatorValue)).clear();
		driver.findElement(By.xpath(LocatorValue)).sendKeys(TestData);
	}
	else if(LocatorType.equalsIgnoreCase("name"))
	{
		driver.findElement(By.name(LocatorValue)).clear();
		driver.findElement(By.name(LocatorValue)).sendKeys(TestData);
	}
}

//method for button,radio,checkbox,link,image click
public static void clickAction(WebDriver driver,String LocatorType, String LocatorValue)
{
	if(LocatorType.equalsIgnoreCase("id"))
	{
		driver.findElement(By.id(LocatorValue)).sendKeys(Keys.ENTER);
	}
	else if(LocatorType.equalsIgnoreCase("xpath"))
	{
		driver.findElement(By.xpath(LocatorValue)).click();
	}
	else if(LocatorType.equalsIgnoreCase("name"))
	{
		driver.findElement(By.name(LocatorValue)).click();
	}
}

//method for 
}
