package base;

import java.io.File;
import java.io.IOException;
import java.time.Duration;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.edge.EdgeDriver;
import org.openqa.selenium.edge.EdgeOptions;


public class Base {
	
	
	public static WebDriver driver;
	
	 static Base b=new Base();
	 
    public static WebDriver openBrowser() throws InvalidFormatException, InterruptedException, IOException{	
    	XSSFWorkbook Credentials = new XSSFWorkbook(new File("C:\\softwares and jars\\IDD_Automation\\test files\\login_details.xlsx"));
		XSSFSheet Sheet = Credentials.getSheetAt(0);
    	System.setProperty("webdriver.edge.driver","C:\\Users\\gadhavek\\Downloads\\msedgedriver.exe");	
		EdgeOptions options=new EdgeOptions();
		options.addArguments("--remote-allow-origins=*");
		WebDriver driver=new EdgeDriver(options);
		driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(10));
		driver.manage().window().maximize();
    	String username =Sheet.getRow(1).getCell(2).toString().trim();
    	String password=Sheet.getRow(1).getCell(3).toString().trim();
    	String client =Sheet.getRow(1).getCell(4).toString().trim();
		//driver.navigate().to("https://www.adpnav.com/");
    	driver.navigate().to(Sheet.getRow(1).getCell(1).toString().trim());
		driver.findElement(By.xpath("//input[@placeholder='Email']")).sendKeys(username);
		driver.findElement(By.xpath("//button[text()='Continue']")).click();
		driver.findElement(By.xpath("//input[@placeholder='Password']")).sendKeys(password);
		driver.findElement(By.xpath("//button[normalize-space()='Login']")).click();
		WebElement element=driver.findElement(By.xpath("//span[text()='"+client+"']"));
		element.click();
		driver.getWindowHandles().forEach(tab->driver.switchTo().window(tab));
        return driver;
    }     
    public static WebDriver openMotif() throws InvalidFormatException, InterruptedException, IOException{	
    	XSSFWorkbook Credentials = new XSSFWorkbook(new File("C:\\softwares and jars\\IDD_Automation\\test files\\login_details.xlsx"));
		XSSFSheet Sheet = Credentials.getSheetAt(0);
    	System.setProperty("webdriver.edge.driver","C:\\Users\\gadhavek\\Downloads\\msedgedriver.exe");	
		EdgeOptions options=new EdgeOptions();
		options.addArguments("--remote-allow-origins=*");
		WebDriver driver=new EdgeDriver(options);
		driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(25));
		driver.manage().window().maximize();		
    	String username =Sheet.getRow(2).getCell(2).toString().trim();
		String password =Sheet.getRow(2).getCell(3).toString().trim();
		String clientid=Sheet.getRow(2).getCell(4).toString().trim();
		driver.navigate().to(Sheet.getRow(2).getCell(1).toString().trim());
		driver.findElement(By.xpath("//input[@id='UserName']")).sendKeys(username);
		driver.findElement(By.xpath("//input[@id='Password']")).sendKeys(password);
		driver.findElement(By.xpath("//input[@value='Log In']")).click();
		WebElement element=driver.findElement(By.xpath("//input[@id='ClientInfoSearchClientField']"));
		element.click();
		element.sendKeys(clientid);
		element.sendKeys(Keys.ARROW_DOWN);
		driver.findElement(By.xpath("//ul[@class='ui-autocomplete ui-front ui-menu ui-widget ui-widget-content']//descendant::li")).click();
		Thread.sleep(1000);
        return driver;
        
    }
    
		
		
	}
    
	
	
		