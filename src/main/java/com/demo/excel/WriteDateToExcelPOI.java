package com.demo.excel;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.concurrent.TimeUnit;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.ui.Select;
import org.testng.annotations.Test;

public class WriteDateToExcelPOI extends DemoBase{
	
	@Test
	public void Register() throws IOException, InterruptedException {

		ArrayList<String> Title = readExcelDate(0);
		ArrayList<String> FirstName = readExcelDate(1);
		ArrayList<String> LastName = readExcelDate(2);
		ArrayList<String> Company = readExcelDate(3);
		
		// ArrayList<String> result= readExcelDate(2);

		for (int i = 1; i <= Title.size(); i++) {
			
			driver.switchTo().frame("mainpanel");
			
			WebElement ele1 =driver.findElement(By.xpath("//*[@id='navmenu']/ul/li[4]/a"));
			Actions a =new Actions(driver);
			a.moveToElement(ele1).build().perform();
			
			driver.findElement(By.xpath("//*[@id='navmenu']/ul/li[4]/ul/li[1]/a")).click();
			Thread.sleep(2000);
			
			Select select =new Select(driver.findElement(By.xpath("//*[@id='contactForm']/table/tbody/"
					+ "tr[2]/td[1]/table/tbody/tr[1]/td[2]/select")));
			select.selectByVisibleText(Title.get(i));
			
			driver.findElement(By.id("first_name")).sendKeys(FirstName.get(i));
			
			driver.findElement(By.id("surname")).sendKeys(LastName.get(i));
			
			driver.findElement(By.name("client_lookup")).sendKeys(Company.get(i));
			
			Thread.sleep(1000);
			
			driver.findElement(By.xpath("//*[@id='contactForm']/table/tbody/tr[1]/td/input[2]")).click();
			
			Thread.sleep(2000);
			
			//driver.findElement(By.xpath("//*[@id='navmenu'/ul/li[1]/a")).click();

			Boolean ispresent=driver.findElements(By.xpath("//*[@id='_mailMergeTemplateSelector']")).size()>0;
			if(ispresent==true)
			{
				//System.out.println("Pass");
				FileInputStream fis = new FileInputStream(
						"/home/sbv6/Desktop/Soumya/eclipse_workspace/"
						+ "demoexcelpoi/src/main/java/com/testdata/dataprovidorfor.xlsx");
				XSSFWorkbook wb = new XSSFWorkbook(fis);
				XSSFSheet s = wb.getSheet("contacts");
				Row row = s.getRow(i);
				Cell cell = row.getCell(4);//changed
				if (cell == null) {
			        cell = row.createCell(4);//changed
			    }
			    cell.setCellType(Cell.CELL_TYPE_STRING);
			    cell.setCellValue("pass");
			    FileOutputStream fos = new FileOutputStream("/home/sbv6/Desktop/Soumya/eclipse_workspace/"
			    		+ "demoexcelpoi/src/main/java/com/testdata/dataprovidorfor.xlsx");
			    wb.write(fos);			 
			    
			  driver.findElement(By.xpath("//*[@id='navmenu']/ul/li[1]/a")).click();			    

			    
			}
			else {
				FileInputStream fis = new FileInputStream(
						"/home/sbv6/Desktop/Soumya/eclipse_workspace/"
						+ "demoexcelpoi/src/main/java/com/testdata/dataprovidorfor.xlsx");
				XSSFWorkbook wb = new XSSFWorkbook(fis);				
				XSSFSheet s = wb.getSheet("contacts");
				Row row = s.getRow(i);
				Cell cell = row.getCell(4);
				if (cell == null) {
			        cell = row.createCell(4);
			    }
			    cell.setCellType(Cell.CELL_TYPE_STRING);
			    cell.setCellValue("Fail");
			    FileOutputStream fos = new FileOutputStream("/home/sbv6/Desktop/Soumya/eclipse_workspace/"
			    		+ "demoexcelpoi/src/main/java/com/testdata/dataprovidorfor.xlsx");
			    wb.write(fos);
			    //driver.navigate().refresh();
			    continue;
			   
				
			}	
			
		}

	}

	public ArrayList<String> readExcelDate(int colNo) throws IOException {
		FileInputStream fis = new FileInputStream(
				"/home/sbv6/Desktop/Soumya/eclipse_workspace/"
				+ "demoexcelpoi/src/main/java/com/testdata/dataprovidorfor.xlsx");
		XSSFWorkbook wb = new XSSFWorkbook(fis);
		XSSFSheet s = wb.getSheet("contacts");
		Row row = s.getRow(1);
		Cell cell = row.getCell(2);
		/*if (cell == null) {
	        cell = row.createCell(2);
	    }
	    cell.setCellType(Cell.CELL_TYPE_STRING);
	    cell.setCellValue("some value");*/
		
		Iterator<Row> rowIterator = s.iterator();
		//rowIterator.next();
		
		ArrayList<String> list = new ArrayList<String>();
		while (rowIterator.hasNext()) {

			list.add(rowIterator.next().getCell(colNo).getStringCellValue());

		}

		System.out.println("List" + list);
		return list;

	}


}
