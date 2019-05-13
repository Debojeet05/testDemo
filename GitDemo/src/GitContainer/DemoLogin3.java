package GitContainer;

import org.testng.annotations.Test;
import org.testng.annotations.BeforeMethod;
import org.testng.annotations.BeforeTest;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.testng.annotations.AfterMethod;
import org.testng.annotations.AfterTest;

public class DemoLogin3 {
	FileInputStream objIS;
	FileOutputStream objOS;
	WebDriver driver;
	//File obj;
	XSSFWorkbook objWB;
	XSSFSheet objSH;
	String row, col, getTEXT, xlmsg;
	String pass= "Test is passed";
	String fail= "Test is failed";
	
	int rowVal;
	int rowCount;
	
	
	
	
	@Test
	public void f() throws InterruptedException, IOException {
		
		System.out.println(rowCount);
		
		
		for( rowVal=1; rowVal< rowCount; rowVal++) {
			
			driver.get("http://demowebshop.tricentis.com/login");	
			
	row= objSH.getRow(rowVal).getCell(0).getStringCellValue();
		col= objSH.getRow(rowVal).getCell(1).getStringCellValue();
		xlmsg= objSH.getRow(rowVal).getCell(2).getStringCellValue();
		
		
		System.out.println("test" +  rowVal + row);
	
		  Thread.sleep(2000);
driver.findElement(By.xpath("//input[@class= 'email']")).sendKeys(row);
		driver.findElement(By.xpath("//*[@id=\"Password\"]")).sendKeys(col);
		driver.findElement(By.xpath("//input[@class= 'button-1 login-button']")).click();
		Thread.sleep(3000);
		getTEXT= driver.findElement(By.cssSelector("a[class='account'][href='/customer/info']")).getText();
		System.out.println(getTEXT);
		  driver.findElement(By.cssSelector("a[class='ico-logout']")).click();
		  Thread.sleep(2000);
		  
		  
		  objOS = new FileOutputStream("C:\\Users\\training_M5.06.15\\Desktop\\Debojeet\\Debojeet.xlsx");
		  objSH.getRow(rowVal).createCell(3).setCellValue(getTEXT);
		  
		  if(xlmsg.equals(getTEXT)) {
			  
			  System.out.println(pass);
			  objSH.getRow(rowVal).createCell(4).setCellValue(pass);
		  }
		  else {
			  
			  System.out.println(fail);
			  objSH.getRow(rowVal).createCell(4).setCellValue(fail);
			  
			  
		  }
			  
		   
	
		  objWB.write(objOS);
		
		}
		
	
	}
	
  @BeforeTest
  public void beforetest() throws InvalidFormatException, IOException {
	
	  System.setProperty("webdriver.chrome.driver", "C:\\Users\\training_M5.06.15\\Desktop\\Debojeet\\chromedriver.exe");
		driver= new ChromeDriver();
	
		
		//obj= new File("C:\\Users\\training_M5.06.15\\Desktop\\Debojeet\\Debojeet.xlsx");
		objIS= new FileInputStream("C:\\Users\\training_M5.06.15\\Desktop\\Debojeet\\Debojeet.xlsx");
		
		objWB= new XSSFWorkbook(objIS);
		objSH= objWB.getSheet("Creds");
		rowCount= objSH.getPhysicalNumberOfRows();
		
		
		
		
  }

  @AfterTest
  public void aftertest() throws IOException {
	  
	
	
	  
	  driver.close();
	  objWB.close();
  }

}
