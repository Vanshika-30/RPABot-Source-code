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
		
		String dbname;
		String bduname;
		String dbpsw;
		
		if(args.length == 3) {
			
			dbname = args[0];
			bduname = args[1];
			dbpsw = args[2];
			
			// Initializing required variables
			WebDriver driver = null;
			String browser = "chrome";
			String url = "jdbc:mysql://localhost:3306/" + dbname;
			String uname = bduname;
			String pass = dbpsw;
			String path_excel = "Questions.xlsx";
			String url_form = "https://rpabot.netlify.app";
			
			// Setting the web driver			
			if(browser.equalsIgnoreCase("firefox")){
					System.setProperty("webdriver.gecko.driver", ".\\geckodriver.exe");
					driver = new FirefoxDriver();
				}
			else if(browser.equalsIgnoreCase("chrome")){
					System.setProperty("webdriver.chrome.driver","chromedriver");
					driver = new ChromeDriver();
				}
			else if(browser.equalsIgnoreCase("Edge")){
					System.setProperty("webdriver.edge.driver",".\\MicrosoftWebDriver.exe");
					driver = new EdgeDriver();
				}
			else if(browser.equalsIgnoreCase("safari")){
					System.setProperty("webdriver.safari.driver", "SafariDriver.safariextz");
					driver = new SafariDriver();
				}

			else{
					throw new Exception("Browser is not correct");
				}
				driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
			
//			driver.manage().window().maximize();
			driver.get(url_form);
			Connection myConn = null;
			
			try{				
				myConn = DriverManager.getConnection(url, uname, pass);
				Statement myStmt = myConn.createStatement();
				
				myStmt.executeUpdate("create table if not exists ques (id int not null auto_increment primary key, ques text, op1 varchar(255), op2 varchar(255), op3 varchar(255), op4 varchar(255), ans varchar(10))");
			} catch (SQLException e1) {
				e1.printStackTrace();
			}
							
			try{

				FileInputStream file = new FileInputStream(path_excel);
				
				@SuppressWarnings("resource")
				XSSFWorkbook wb = new XSSFWorkbook(file);
				
				XSSFSheet sheet = wb.getSheet("Sheet1");
				
				int rowCount = sheet.getLastRowNum();
				
				for(int row = 1; row <= rowCount; row++) 
				{
					XSSFRow currRow = sheet.getRow(row);
					
					String ques = currRow.getCell(1).getStringCellValue();
					String op1 = currRow.getCell(2).getStringCellValue();
					String op2 = currRow.getCell(3).getStringCellValue();
					String op3 = currRow.getCell(4).getStringCellValue();
					String op4 = currRow.getCell(5).getStringCellValue();
					String ans = currRow.getCell(6).getStringCellValue();
									
					// Upload Process					
					if(!ques.isEmpty() && !op1.isEmpty() && !ans.isEmpty() && !op2.isEmpty()) {
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
						driver.manage().timeouts().implicitlyWait(10, TimeUnit.MINUTES);
						
						// Saving to database
						if (driver.findElement(By.tagName("p")).getText().equals("Submitted Successfully")) {
							driver.findElement(By.tagName("a")).click();
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
						}
						
						 else if (driver.findElement(By.tagName("p")).getText().equals("Modified and submitted Successfully"))
						 {
							 
								String ques_mod = driver.findElement(By.id("v1")).getText();
								String op1_mod = driver.findElement(By.id("v2")).getText();
								String op2_mod= driver.findElement(By.id("v3")).getText();
								String op3_mod = driver.findElement(By.id("v4")).getText();
								String op4_mod = driver.findElement(By.id("v5")).getText();
								String ans_mod = driver.findElement(By.id("v6")).getText();
								
							try {								
								Statement myStmt = myConn.createStatement();
								Statement myStmt1 = myConn.createStatement();
								String check = "select * from ques where ques in ('" + ques_mod + "')";
								ResultSet rs = myStmt1.executeQuery(check);
								if(!rs.next()) {
									String query = "insert into ques(ques,op1,op2,op3,op4,ans) values ('"+ques_mod+"','"+op1_mod+"','"+op2_mod+"','"+op3_mod+"','"+op4_mod+"','"+ans_mod+"')" ;
									
									myStmt.executeUpdate(query);
								}												
							}catch(Exception exc) {
								exc.printStackTrace();
							}							
							driver.findElement(By.tagName("a")).click();
							
						}						
						 else if(driver.findElement(By.tagName("p")).getText().equals("Deleted Successfully")) {
								 driver.findElement(By.tagName("a")).click();
								 
									try {
										Statement myStmt = myConn.createStatement();
										Statement myStmt1 = myConn.createStatement();
										String check = "select * from ques where ques in ('" + ques + "')";
										ResultSet rs = myStmt1.executeQuery(check);
										if(rs.next()) {
											String query = "DELETE FROM ques WHERE  ques like \""+ ques.toString() + "\"";
											myStmt.executeUpdate(query);
										}
														
									}catch(Exception exc) {
										exc.printStackTrace();
									
									}
						 }					
					}
				}	
				file.close();
				
			}catch (Exception e) {
				e.printStackTrace();
			}
			myConn.close();
			driver.quit();
			
		}
		else {
			System.out.println("Please give name of database, database username and password in the same order.");
		}

	}

}
