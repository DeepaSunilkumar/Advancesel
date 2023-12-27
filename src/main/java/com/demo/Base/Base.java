package com.demo.Base;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Properties;

import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;

public class Base {
	public WebDriver driver;
	public Properties prop;
	
	public void init_Driver(String browser) {
		
		if(browser.equalsIgnoreCase("chrome")) {
			driver = new ChromeDriver();
			
		}
	}
	public Properties init_Property() {
		prop = new Properties();
		FileInputStream fis = null;
		try {
			fis = new FileInputStream("C:\\Users\\Deepa\\eclipse-workspace\\DeepaSunil\\src\\main\\java\\com\\demo\\Excel\\Keyword driven testing.xls");
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		try {
			prop.load(fis);
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		return prop;
	}

}
