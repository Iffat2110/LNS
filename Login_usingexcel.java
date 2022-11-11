package LNS;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Iterator;
import java.util.logging.LogManager;

import org.apache.logging.log4j.spi.LoggerContext;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;

public class Login_usingexcel {
	WebDriver driver = new ChromeDriver();
	static {
		System.setProperty("webdriver.chrome.driver",
				"C:\\Users\\Imma\\eclipse-workspace\\LeadingNetworking Systems\\Drivers\\chromedriver.exe");
		System.setProperty("webdriver.chrome.verboseLogging", "true");
		System.setProperty("log4j.configurationFile", "./path_to_the_log4j2_config_file/log4j2.xml");

		// LoggerContext context = (org.apache.logging.log4j.core.LoggerContext)
		// LogManager.getContext(false);
		// File file = new File("path/to/a/different/log4j2.xml");

		// this will force a reconfiguration
		// context.setConfigLocation(file.toURI());

	}

	/**
	 * @param args
	 * @throws EncryptedDocumentException
	 * @throws IOException
	 * @throws NullPointerException
	 * @throws Exception
	 */
	public static void main(String[] args)
			throws EncryptedDocumentException, IOException, NullPointerException, Exception {
		// TODO Auto-generated method stub
		WebDriver driver = new ChromeDriver();
		driver.get("https://cmsweb.m-staging.in/LNS_Support_Sales_Purchase_Testing/");
		driver.manage().window().maximize();
		FileInputStream fis = new FileInputStream(
				"C://Users//Imma//Desktop//Leading_Networking_System/logging_file.xlsx");
		Workbook wb = WorkbookFactory.create(fis);
		String data = wb.getSheet("Sheet1").getRow(1).getCell(1).toString();
		WebElement email = driver.findElement(By.name("email"));
		email.sendKeys(data);
		// System.out.println(data);

		String data1 = wb.getSheet("Sheet1").getRow(1).getCell(2).toString();
		WebElement password = driver.findElement(By.name("password"));
		password.sendKeys(data1);
		// System.out.println(data1);
		Thread.sleep(2000);
		WebElement checkbox = driver.findElement(By.name("remember"));
		checkbox.click();
		Thread.sleep(2000);
		WebElement SignIn = driver.findElement(By.name("submit"));
		SignIn.click();
		/*
		 * driver.findElement(By.xpath(
		 * "/html[1]/body[1]/div[1]/aside[1]/section[1]/ul[1]/li[2]/a[1]")).click();
		 * Thread.sleep(2000); driver.findElement(By.xpath(
		 * "/html[1]/body[1]/div[1]/aside[1]/section[1]/ul[1]/li[2]/ul[1]/li[1]/a[1]"))
		 * .click(); Thread.sleep(2000);
		 */
		// Add Branch on empty
		driver.findElement(By.xpath("//span[contains(text(),'Master')]")).click();
		
		Thread.sleep(2000);
		driver.findElement(By.xpath("//body/div[2]/aside[1]/section[1]/ul[1]/li[2]/ul[1]/li[1]/a[1]")).click();
		Thread.sleep(2000);
		// driver.close();
		// To check if the branch accepts the duplicate name or no
		Sheet sheet = wb.getSheetAt(1);
		/*
		 * Row row = sheet.getRow(0); int colNum = row.getLastCellNum();
		 * System.out.println("Total Number of Columns in Excel" +colNum ); int rowNum
		 * =sheet.getLastRowNum()+1; System.out.println("Total Number of Rows  in Excel"
		 * +rowNum);
		 * 
		 * driver.close(); for (int j = 0; j < row.getLastCellNum(); j++) { for (int i =
		 * 0; i < sheet.rowNum ; i++) {
		 * 
		 * 
		 * }
		 * 
		 * 
		 * }
		 * 
		 * String data2=wb.getSheet("Branch").getRow(1).getCell(1).toString();
		 */
		Iterator <Row> rowIt=sheet.iterator();
		while(rowIt.hasNext()) 
		{
			
			Row row=rowIt.next();
			Iterator <Cell> cellIterator =row.cellIterator();
			while(cellIterator.hasNext())
			{
				Cell cell =cellIterator.next();
				String st=cell.toString();
				/* System.out.println(st); */
				driver.findElement(By.name("branch_name")).sendKeys(st);

				driver.findElement(By.name("submit_btn")).click();
				Thread.sleep(1000);
				if(driver.getPageSource().contains("* This branch is already exist"))
				{
					driver.findElement(By.xpath("//body/div[2]/div[1]/section[1]/div[1]/div[1]/div[1]/div[1]/form[1]/div[2]/a[1]/button[1]")).click();
				}
				/* driver.findElement(By.name("submit_btn")).click(); */
			}
			
		}

		driver.close();
	}

}
