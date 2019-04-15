package com.demo.excel;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.testng.annotations.BeforeMethod;
import org.testng.annotations.Test;

public class DemoBase {
	
	

	public static WebDriver driver;
	
	@BeforeMethod
	public void start() throws InterruptedException{
		System.setProperty("webdriver.chrome.driver","//home//sbv6//Downloads//chromedriver");		
		driver =  new ChromeDriver(); 
		driver.get("https://www.freecrm.com/index.html");
		driver.manage().window().maximize();
		driver.findElement(By.name("username")).sendKeys("soumya456");
		driver.findElement(By.name("password")).sendKeys("$oumya@12");
		Thread.sleep(3000);
		driver.findElement(By.xpath("//input[@type='submit']")).click();
		
	}

}
