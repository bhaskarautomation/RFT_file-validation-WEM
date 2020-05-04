package Practice1;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;

public class Practice1 {

	public static void main(String[] args) throws IOException  {
		// TODO Auto-generated method stub

		XSSFSheet sheet;
		File src = new File("C:\\Users\\AG22518\\Desktop\\bhavya\\Practice1.xlsx");
		
		FileInputStream fis = new FileInputStream(src);
		
		XSSFWorkbook wb = new XSSFWorkbook(fis);
			
		sheet = wb.getSheet("Cred");
		
		Practice1 b = new Practice1();
		
		String url = b.ReadfromExcel("URL", 0, 1);
							
		System.setProperty("webdriver.chrome.driver","C:\\apache\\chromedriver\\chromedriver.exe");
		
		WebDriver driver = new ChromeDriver();
		
		driver.get(url);
		String username = b.ReadfromExcel("Cred", 0, 1);
		String password = b.ReadfromExcel("Cred", 1, 1);
		
		driver.findElement(By.xpath("//*[@id='USER']")).sendKeys(username);
		driver.findElement(By.xpath("//*[@id='PASSWORD']")).sendKeys(password);
		driver.findElement(By.xpath("//*[@id='signInButton']")).click();
		
		sheet.getRow(0).createCell(2).setCellValue("pass - login successfull");
		
		
				
}
	public static String ReadfromExcel(String sheet,int row,int cell) throws IOException{
		
		FileInputStream fis = new FileInputStream("C:\\Users\\AG22518\\Desktop\\bhavya\\Practice1.xlsx");
		
		XSSFWorkbook wb = new XSSFWorkbook(fis);
		XSSFSheet sheet1 = wb.getSheet(sheet);
		XSSFRow row1 = sheet1.getRow(row);
		XSSFCell cell1 = row1.getCell(cell);
		String data = cell1.getStringCellValue();	
		
		return data;
		
		
	}
	
	
	
	

}