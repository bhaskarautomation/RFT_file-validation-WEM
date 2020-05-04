

package wem;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;

public class webvalidation {

	public static void main(String[] args) throws IOException {
		
		XSSFSheet Sheet;

        File src = new File("C:\\Users\\AG22518\\Desktop\\bhavya\\wemdata1.xlsx");

        FileInputStream fis = new FileInputStream(src);

        XSSFWorkbook wb = new XSSFWorkbook(fis);

        Sheet = wb.getSheet("Input");
		
        webvalidation b = new webvalidation();
        
        String Userid = b.Readfromexcel("Credentials",0,1);
        String password = b.Readfromexcel("Credentials",1,1);
        
        String URL = b.Readfromexcel("URL",0,1);
               
        
		//String URL = "https://wem.uat.va.anthem.com";
		
		System.setProperty("webdriver.chrome.driver","C:\\apache\\chromedriver\\chromedriver.exe");
		
		WebDriver driver = new ChromeDriver();
		
		driver.get(URL);
		
		driver.findElement(By.id("USER")).sendKeys(Userid);
		
		driver.findElement(By.id("PASSWORD")).sendKeys(password);
		
		driver.findElement(By.id("signInButton")).click();
		
		
		for(int i=1;i<=Sheet.getLastRowNum();i++)
			
		{
	
		DataFormatter Formatter = new DataFormatter();

        String SSN = Formatter.formatCellValue(Sheet.getRow(i).getCell(0));
        
          
		driver.findElement(By.id("ssn")).sendKeys(SSN);
		
		driver.findElement(By.id("button")).click();
		
		driver.findElement(By.xpath("//*[@id='applicationTable']/tbody/tr[1]/td[2]")).click();
		
		//*[@id="applicationTable"]/tbody/tr[1]/td[2]
		
	String state ="/html/body/div[1]/div[3]/div[2]/div[1]/div[1]/table/tbody/tr[1]/td[4]/span";
	String stateCode = driver.findElement(By.xpath(state)).getText();
	Sheet.getRow(i).createCell(1).setCellValue(stateCode);
	String wemstatus = driver.findElement(By.xpath("/html/body/div[1]/div[3]/div[2]/div[1]/div[1]/table/tbody/tr[3]/td[2]")).getText();
		
		System.out.println(wemstatus);
		
		Sheet.getRow(i).createCell(2).setCellValue(wemstatus);
		
		String EXsubid= driver.findElement(By.xpath("//*[@id='ExchangeAssignedSubscriberIdLabel']")).getText();
	
		System.out.println(EXsubid);
		
		Sheet.getRow(i).createCell(3).setCellValue(EXsubid);

	        
	     
	        
	        //*[@id='SearchEnrollments']
	        
	        
		driver.findElement(By.xpath("//*[@id='transactionHistory1']")).click();
		
		//*[@id="applicationTable"]/tbody/tr[13]/td[9]/label
		
		
		WebElement htmltable= driver.findElement(By.xpath("//*[@id='applicationTable']/tbody"));
		
		List <WebElement> rowcount = htmltable.findElements(By.tagName("tr"));
		
		//System.out.println(rowcount.size());
		int count = rowcount.size();
		
		System.out.println(count);
		
		String Timestam = "//*[@id='applicationTable']/tbody/tr["+count+"]/td[2]";
		
		String desc = "//*[@id='applicationTable']/tbody/tr["+count+"]/td[9]";
		
		String Timestamp= driver.findElement(By.xpath(Timestam)).getText();
		
		System.out.println(Timestamp);
		
		String Description= driver.findElement(By.xpath(desc)).getText();
		
		System.out.println(Description);
		
		//*[@id='applicationTable']/tbody/tr[11]/td[2]
		
		
		Sheet.getRow(i).createCell(4).setCellValue(Timestamp);
		
		
		
		Sheet.getRow(i).createCell(5).setCellValue(Description);
		
		String flag1 ="/html/body/div[1]/div[3]/div[2]/div[1]/div[1]/table/tbody/tr[10]/td[2]";
		
		String Hflag = driver.findElement(By.xpath(flag1)).getText();
		
		Sheet.getRow(i).createCell(6).setCellValue(Hflag);
		
		String Eind ="/html/body/div[1]/div[3]/div[2]/div[1]/div[1]/table/tbody/tr[9]/td[4]";
	
		String EnrActInd =	driver.findElement(By.xpath(Eind)).getText();
		
		Sheet.getRow(i).createCell(7).setCellValue(EnrActInd);
		
		
		String SeqNum1 = "/html/body/div[1]/div[3]/div[2]/div[1]/div[1]/table/tbody/tr[8]/td[2]";
		String SeqNum = driver.findElement(By.xpath(SeqNum1)).getText();
		String appId = "/html/body/div[1]/div[3]/div[2]/div[1]/div[1]/table/tbody/tr[2]/td[4]";
		String AppliID = driver.findElement(By.xpath(appId)).getText();
		
		Sheet.getRow(i).createCell(8).setCellValue(SeqNum);
		Sheet.getRow(i).createCell(9).setCellValue(AppliID);
		
		String pendr="/html/body/div[1]/div[3]/div[2]/div[1]/div[1]/table/tbody/tr[5]/td[4]";
				String PendReas = driver.findElement(By.xpath(pendr)).getText();
		Sheet.getRow(i).createCell(10).setCellValue(PendReas);
		
				driver.findElement(By.xpath("//*[@id='SearchEnrollments']")).click();
		
		
	      FileOutputStream fileOutput = new FileOutputStream(src);

	        // finally write content

	        wb.write(fileOutput);

	        // close the file
	        fileOutput.close();
	        
		
		//*[@id='applicationTable']/tbody
		
		
		
		//*[@id="transactionHistory1"]
		
		}
	}

    public static String Readfromexcel(String sheet, int row, int cell) throws IOException{
   	 
   	 
   	    FileInputStream fis = new FileInputStream("C:\\Users\\AG22518\\Desktop\\bhavya\\wemdata1.xlsx");
        
        XSSFWorkbook wb = new XSSFWorkbook(fis);
        
        XSSFSheet sheet1 = wb.getSheet(sheet);
        
        XSSFRow row1 = sheet1.getRow(row);
         
        XSSFCell cell1 = row1.getCell(cell);
        
        String data = cell1.getStringCellValue();
		         
        return data; 
   	 
   	 
    }
	
}
