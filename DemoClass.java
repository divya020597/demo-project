package org.demo;

import org.openqa.selenium.Alert;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.interactions.Actions;

public class DemoClass {
public static void main(String[] args) {
	System.setProperty("webdriver.chrome.driver", "C:\\Users\\kavi Divi\\eclipse-workspace\\Selenium\\Driver\\chromedriver.exe");
	WebDriver driver = new ChromeDriver();
	driver.get("http://greenstech.in/selenium-course-content.html");
	Actions a = new Actions(driver);
	 Alert alert = driver.switchTo().alert();
	 TakesScreenshot tk = (TakesScreenshot)driver;
	
		
}
}
