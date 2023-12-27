package com.demo.ExecutionEngine;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Properties;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.sl.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;

import com.demo.Base.Base;

public class Execution {
	
	public WebDriver driver;
	public Properties prop;
	public Workbook book;
	public org.apache.poi.ss.usermodel.Sheet sheet;
	public Base base;
	public final String filepath="C:\\Users\\Deepa\\eclipse-workspace\\DeepaSunil\\src\\main\\java\\com\\demo\\Excel\\Keyword driven testing.xls";
	
	public void startEngine(String sheetname) {
		String locatortype = null;
		String locatorvalue = null;
		FileInputStream fis = null;
		WebElement element = null;;
		
		try {
			fis = new FileInputStream(filepath);
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		try {
			book = WorkbookFactory.create(fis);
		} catch (EncryptedDocumentException e) {
			// TODO Auto-generated catch block
            e.printStackTrace();
			
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		
		sheet = book.getSheet(sheetname);
		int j = 0;
		for(int i=0;i<sheet.getLastRowNum();i++) {
			try {
			
			String locator = sheet.getRow(i+1).getCell(j+1).toString().trim();
			
			if(!locator.equalsIgnoreCase("na")); {
				
				locatortype = locator.split("=")[0].toString().trim();
				locatorvalue = locator.split("=")[1].toString().trim();
			}
			
			String Action = sheet.getRow(i+1).getCell(i+2).toString().trim();
			String Value = sheet.getRow(i+1).getCell(j+3).toString().trim();
			
			switch (Action) {
			case "open browser":
			base = new Base();
			prop = base.init_Property();
			if(Value.isEmpty()||Value.equalsIgnoreCase("Na")) {
				base.init_Driver(prop.getProperty("browser"));
			}
			else {
				base.init_Driver(Value);
			}
			break;
			
			
			case "enter url":
				if(Value.isEmpty()||Value.equalsIgnoreCase("NA")) {
					driver.get(prop.getProperty("url"));
				}
				else {
					driver.get(Value);
				}
				break;
				
				
				
			case "quit":
				driver.quit();
				break;
			}
				
				
				switch (locatortype) {
				case "id":
					 element = driver.findElement(By.id(locatorvalue));
					
					if(Action.equals("click")) {
						element.click();
					}
					else if(Action.equals("sendkeys")) {
						element.clear();
						element.sendKeys(Value);
					}
					break;
					
				case "linkText":
					
					element = driver.findElement(By.linkText(locatorvalue));
					
					if(Action.equalsIgnoreCase("click")) {
						element.click();
						
					}
					else if(Action.equalsIgnoreCase("sendkeys")) {
						element.clear();
						element.sendKeys(Value);
						
					}
					break;
					
				case "cssSelector":
					
					element = driver.findElement(By.cssSelector(locatorvalue));
					
					if(Action.equalsIgnoreCase("click")) {
						element.click();
					}
					
					else if(Action.equalsIgnoreCase("sendkeys")) {
						element.clear();
						element.sendKeys(Value);
					}
					break;
					
					
					
					
				
				
			default:
			break;
			} }
			catch(Exception e) {
				
			}
		}
		
		  
		 
	}

}
