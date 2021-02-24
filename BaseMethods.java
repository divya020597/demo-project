package org.demo;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;

public class BaseMethods {
public static WebDriver driver;
	
	//method1
	public static  WebDriver launchbrowser() {
		System.setProperty("webdriver.chrome.driver","C:\\Users\\kavi Divi\\eclipse-workspace\\Maven\\drivers\\chromedriver.exe");
		driver = new ChromeDriver();
		return driver;
	}
	//m2
	public static void launchurl(String url) {
		driver.get(url);
	}
	//m3
	public static void filltextbox(WebElement e,String st) {
	 e.sendKeys(st);
}
	public static String getAttribute(WebElement e) {
		return e.getAttribute("value");
		
}
	
	public static void btnclick(WebElement e) {
		e.click();
	}
	public static void maxi() {
     driver.manage().window().maximize();
	}
	public static String gettitle() {
		return driver.getTitle();
	}
	public static String gettext(WebElement e){
		return e.getText();
		}
	public static  void dataupdate(String sheetname , int rowno,int cellno, String currentvalue,String updatevalue) throws Exception {
		File loc = new File(
				"C:\\Users\\kavi Divi\\eclipse-workspace\\Maven\\src\\test\\resources\\xmlfile\\Book1.xlsx");
		// get the file
		FileInputStream st = new FileInputStream(loc);
		Workbook wk = new XSSFWorkbook(st);
		Sheet sheet = wk.getSheet(sheetname);
		Row row = sheet.getRow(rowno);
		Cell cell = row.getCell(cellno);
		String stringCellValue = cell.getStringCellValue();
		if(stringCellValue.equals(currentvalue)) {
			cell.setCellValue(updatevalue);
			
		}
		FileOutputStream o=new FileOutputStream(loc);
		wk.write(o);
		
	}

}
