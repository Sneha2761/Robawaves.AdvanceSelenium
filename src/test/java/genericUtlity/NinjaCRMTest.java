package genericUtlity;

import java.io.FileInputStream;
import java.io.IOException;
import java.time.Duration;
import java.util.Properties;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.edge.EdgeDriver;

public class NinjaCRMTest {

	public static void main(String[] args) throws EncryptedDocumentException, IOException {
		FileInputStream fs=new FileInputStream(".\\src\\test\\resources\\nynga.properties");
		Properties pro = new Properties();
		pro.load(fs);
		String url = pro.getProperty("url");
		String username = pro.getProperty("username");
		String pass = pro.getProperty("password");
		
		FileInputStream fis=new FileInputStream(".\\src\\test\\resources\\example.xlsx");
		Workbook wb = WorkbookFactory.create(fis);
		Sheet sheet = wb.getSheet("example");
		Row row = sheet.getRow(1);
		Cell cell = row.getCell(2);
		double targetSizeNum = sheet.getRow(1).getCell(3).getNumericCellValue();
		String target_size = String.valueOf((int) targetSizeNum);
		String value1 = cell.getStringCellValue();
		
		
		
		WebDriver driver = new EdgeDriver();
		driver.manage().window().maximize();
		driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(10));
		driver.get(url);
		driver.findElement(By.id("username")).sendKeys(username);
		driver.findElement(By.id("inputPassword")).sendKeys(pass);
		driver.findElement(By.xpath("//button[@type='submit']")).click();
		driver.findElement(By.cssSelector("[class='btn btn-info']")).click();
		driver.findElement(By.cssSelector("[name='campaignName']")).sendKeys(value1);
		WebElement ts = driver.findElement(By.cssSelector("[type='number']"));
		ts.clear();
		ts.sendKeys(target_size);
		driver.findElement(By.cssSelector("[type='submit']")).click();
		

	}

}
