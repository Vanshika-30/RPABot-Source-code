package com.rpabot;

import java.io.FileInputStream;
import java.sql.*;
import java.util.concurrent.TimeUnit;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.edge.EdgeDriver;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.safari.SafariDriver;


public class UploadQues {

	public static void main(String[] args) throws Exception{
		
		// Initializing required variables
		WebDriver driver = null;
		String browser = "chrome";
		String url = "jdbc:mysql://localhost:3306/question";
		String uname = "root";
		String pass = "Learn#30";
		String path_excel = "Questions.xlsx";
		String url_form = "file:///Users/vanshikajain/eclipse-workspace/RPABot/index.html";
		
		// Setting the web driver
		if(browser.equalsIgnoreCase("firefox")){
			//create firefox instance
				System.setProperty("webdriver.gecko.driver", ".\\geckodriver.exe");
				driver = new FirefoxDriver();
			}
		else if(browser.equalsIgnoreCase("chrome")){
				//set path to chromedriver.exe
				System.setProperty("webdriver.chrome.driver","chromedriver");
				//create chrome instance
				driver = new ChromeDriver();
			}
			//Check if parameter passed as 'Edge'
		else if(browser.equalsIgnoreCase("Edge")){
						//set path to Edge.exe
				System.setProperty("webdriver.edge.driver",".\\MicrosoftWebDriver.exe");
						//create Edge instance
				driver = new EdgeDriver();
			}
		else if(browser.equalsIgnoreCase("safari")){
				System.setProperty("webdriver.safari.driver", "SafariDriver.safariextz");
				driver = new SafariDriver();
			}

		else{
				//If no browser passed throw exception
				throw new Exception("Browser is not correct");
			}
			driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
		
			
		driver.get(url_form);
		Connection myConn = null;
		
		try{
			
			myConn = DriverManager.getConnection(url, uname, pass);
			Statement myStmt = myConn.createStatement();
			
			myStmt.executeUpdate("create table if not exists ques (id int not null auto_increment primary key, ques text, op1 varchar(255), op2 varchar(255), op3 varchar(255), op4 varchar(255), ans varchar(2))");
		} catch (SQLException e1) {
			e1.printStackTrace();
		}
			
		
		try{

			FileInputStream file = new FileInputStream(path_excel);
			
			@SuppressWarnings("resource")
			XSSFWorkbook wb = new XSSFWorkbook(file);
			
			XSSFSheet sheet = wb.getSheet("Sheet1");
			
			int rowCount = sheet.getLastRowNum();
//			System.out.println(rowCount);
			
			for(int row = 1; row <= rowCount; row++) 
			{
//				driver.findElement(By.linkText("Upload Another Question")).click();
//				Thread.sleep(3000);
				XSSFRow currRow = sheet.getRow(row);
				
				String ques = currRow.getCell(1).getStringCellValue();
				String op1 = currRow.getCell(2).getStringCellValue();
				String op2 = currRow.getCell(3).getStringCellValue();
				String op3 = currRow.getCell(4).getStringCellValue();
				String op4 = currRow.getCell(5).getStringCellValue();
				String ans = currRow.getCell(6).getStringCellValue();
								
				// Upload Process
				
				// Cleaning the fields
				driver.findElement(By.name("ques")).clear();
				driver.findElement(By.name("op1")).clear();
				driver.findElement(By.name("op2")).clear();
				driver.findElement(By.name("op3")).clear();
				driver.findElement(By.name("op4")).clear();
				driver.findElement(By.name("ans")).clear();
				
				
				// Uploading data
				driver.findElement(By.name("ques")).sendKeys(ques);
				driver.findElement(By.name("op1")).sendKeys(op1);
				driver.findElement(By.name("op2")).sendKeys(op2);
				driver.findElement(By.name("op3")).sendKeys(op3);
				driver.findElement(By.name("op4")).sendKeys(op4);
				driver.findElement(By.name("ans")).sendKeys(ans);
				
				// Submitting the form
//				Thread.sleep(2000);
//				driver.findElement(By.linkText("SUBMIT")).click();
				
				
				// Saving to database
				
				try {
					Statement myStmt = myConn.createStatement();
					Statement myStmt1 = myConn.createStatement();
					String check = "select * from ques where ques in ('" + ques + "')";
					ResultSet rs = myStmt1.executeQuery(check);
					if(!rs.next()) {
						String query = "insert into ques(ques,op1,op2,op3,op4,ans) values ('"+ques+"','"+ op1+"','"+op2+"','"+op3+"','"+op4+"','"+ans+"')" ;
						
						myStmt.executeUpdate(query);

					}
										
				}catch(Exception exc) {
					exc.printStackTrace();
				}
//				Thread.sleep(2000);	
			}
			
			file.close();
			
		}catch (Exception e) {
//			System.out.println("1");
			e.printStackTrace();
		}
		myConn.close();
		driver.quit();
	}

}
