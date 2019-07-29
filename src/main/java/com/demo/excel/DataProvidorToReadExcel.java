package com.demo.excel;

import java.io.FileInputStream;

import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Set;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.openqa.selenium.By;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.ui.Select;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

import utils.TestUtil;

public class DataProvidorToReadExcel extends DemoBase {
	
	@DataProvider
	public Object[][] getRegister(){
		Object data[][] = TestUtil.getTestData("contacts");//"SHEET NAME"
		return data;
		
	}
	
	@Test(dataProvider="getRegister")
	public void login(String title,String fn,String ln,String Company) throws InterruptedException {
		Thread.sleep(2000);
		
		driver.switchTo().frame("mainpanel");
		
		WebElement ele1 =driver.findElement(By.xpath("//*[@id='navmenu']/ul/li[4]/a"));
		Actions a =new Actions(driver);
		a.moveToElement(ele1).build().perform();
		
		driver.findElement(By.xpath("//*[@id='navmenu']/ul/li[4]/ul/li[1]/a")).click();
		Thread.sleep(2000);
		
		Select select =new Select(driver.findElement(By.xpath("//*[@id='contactForm']/table/tbody/"
				+ "tr[2]/td[1]/table/tbody/tr[1]/td[2]/select")));
		select.selectByVisibleText(title);
		
		driver.findElement(By.id("first_name")).sendKeys(fn);
		
		driver.findElement(By.id("surname")).sendKeys(ln);
		
		driver.findElement(By.name("client_lookup")).sendKeys(Company);
		
		driver.findElement(By.xpath("//*[@id='contactForm']/table/tbody/tr[1]/td/input[2]")).click();
		
	
	}

}
