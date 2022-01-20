package org.che;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Date;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.Select;

import io.github.bonigarcia.wdm.WebDriverManager;

public class Task {
	public static void main(String[] args) throws IOException {
		WebDriverManager.chromedriver().setup();
		WebDriver driver= new ChromeDriver();
		driver.get("http://demo.automationtesting.in/Register.html");
		WebElement element = driver.findElement(By.xpath("//select[@type='text']"));
		Select select = new Select(element);
		List<WebElement> options = select.getOptions();
		File file = new File("C:\\Users\\DELL\\eclipse-workspace\\Lxs\\lib\\Hello.xlsx");
		Workbook workbook = new XSSFWorkbook();
		Sheet sheet = workbook.createSheet("Test");
		for (int i = 0; i < options.size(); i++) {
		WebElement element2 = options.get(i);
		String text = element2.getText();
		System.out.println(text);
		
		Row row = sheet.createRow(i);
		Cell cell = row.createCell(0);
		cell.setCellValue(text);
		}
 			FileOutputStream o = new FileOutputStream(file);
 			workbook.write(o);
			
			
			driver.quit();
			}

		



	
}





