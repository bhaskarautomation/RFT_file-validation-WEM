
import java.awt.AWTException;
import java.awt.Robot;
import java.awt.event.KeyEvent;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import resources.Script1Helper;

import com.rational.test.ft.*;
import com.rational.test.ft.object.interfaces.*;
import com.rational.test.ft.object.interfaces.SAP.*;
import com.rational.test.ft.object.interfaces.WPF.*;
import com.rational.test.ft.object.interfaces.dojo.*;
import com.rational.test.ft.object.interfaces.siebel.*;
import com.rational.test.ft.object.interfaces.flex.*;
import com.rational.test.ft.object.interfaces.generichtmlsubdomain.*;
import com.rational.test.ft.script.*;
import com.rational.test.ft.value.*;
import com.rational.test.ft.vp.*;
import com.ibm.rational.test.ft.object.interfaces.sapwebportal.*;
/**
 * Description   : Functional Test Script
 * @author AG22518
 */
public class Script1 extends Script1Helper
{
	/**
	 * Script Name   : <b>Script1</b>
	 * Generated     : <b>Sep 16, 2019 2:44:23 PM</b>
	 * Description   : Functional Test Script
	 * Original Host : WinNT Version 10.0  Build 16299 ()
	 * 
	 * @since  2019/09/16
	 * @author AG22518
	 * @throws IOException 
	 * @throws AWTException 
	 */
	public void testMain(Object[] args) throws IOException, AWTException 
	{
	
	
	
		XSSFSheet Sheet;

        File src = new File("C:\\Users\\AG22518\\Desktop\\bhavya\\input.xlsx");

        FileInputStream fis = new FileInputStream(src);

        XSSFWorkbook wb = new XSSFWorkbook(fis);

        Sheet = wb.getSheet("Input");

        startApp("Extension for Terminal Applications"); 
        
        sleep(6);
        
        Property[] properties = new Property[2];
        properties[0] = new Property(".class", "com.ibm.terminal.tester.gui.misc.AccessibleTextField");
        properties[1]= new Property(".classIndex", "0");
        sleep(0.3);
        TestObject[] text = find(atDescendant(properties));
        TextGuiSubitemTestObject IPaddress = ((TextGuiSubitemTestObject)text[0]);
        IPaddress.waitForExistence();
        IPaddress.setText("30.130.200.57");
        
        
        Property b1 = new Property(".class","javax.swing.JButton");
        Property b2 = new Property(".classIndex","4");
        Property[] propertie = {b1,b2};
        TestObject[] a = find(atDescendant(propertie));
        ((GuiTestObject)a[0]).click();
        sleep(2);
  
  
        Script1 b = new Script1();
        String Userid = b.Readfromexcel("Credentials",0,1);
        String password = b.Readfromexcel("Credentials",1,1);
        
        String Filename = b.Readfromexcel("File",0,1);
        
       
        
        b.enter_text(23, 48, "sst");
        
        b.enter();
        
        sleep(2);
        
        b.enter_text(14, 20, Userid);
        
        b.enter_text(15, 20, password);
                
        b.enter();
        
        sleep(2);
        
        b.enter_text(23, 15, "t TSOK");
        
        b.enter();
        
        		
		// 
		field_18_20().click(atPoint(5,30));
		ibmExtensionForTerminalBasedAp().inputChars("2");
        
        
//        b.enter_text(18, 26, "2");
//        
        b.enter();           
       
       sleep(2);
        
        b.enter_text(23, 15, "TSOK");
        
        b.enter();
        
        sleep(2);
        
        b.enter();
        
        sleep(2);
                
        b.enter();
        
        sleep(2);
        
        b.enter_text(21, 14, "3.4");
        
        b.enter();
         
        sleep(1);
        
        b.enter_text(9, 24, Filename);
        
        b.enter();
        
        sleep(2);
        
        b.enter_text(7, 2, "v");
        
        b.enter();
        
        b.enter();
        
        sleep(2);
        
//        String input = b.Readfromexcel("Input",0,1);
//        
//        b.enter_text(7, 4, "f "+input);
//        
//        b.enter();
        
		
		//String Field_3_56_text = (String)field_3_56().getProperty("text");
	   
        
        for(int i=1;i<=Sheet.getLastRowNum();i++)

        {
        // 
        DataFormatter Formatter = new DataFormatter();

        String input = Formatter.formatCellValue(Sheet.getRow(i).getCell(0));

        b.enter_text(21, 15, "f "+input);
        
        b.enter();
       
        String dataa = (String)field_3_56().getProperty("text");
        
//        System.out.println(dataa); 
//               
//        if(dataa.contains("*Bottom of data reached*"))
//        	
//        {
//        	b.F5();
//        	
//        	
//      	
//        }
//        
//       
//        String data1 = (String)field_3_56().getProperty("text");	
//        
//        System.out.println(data1);
//        
//      Sheet.getRow(i).createCell(1).setCellValue(data1);
        
        
        
		
		// 
		field_21_15().click(atPoint(11,32));
		ibmExtensionForTerminalBasedAp().inputKeys("m{F7}");
        
        
//        b.enter_text(21, 15, "M");
//        
//              
//        b.F7();
//        
//        sleep(1);
        
      
        Sheet.getRow(i).createCell(1).setCellValue(dataa);
        
        FileOutputStream fileOutput = new FileOutputStream(src);

        // finally write content

        wb.write(fileOutput);

        // close the file
        fileOutput.close();


        }

	}
     public static String Readfromexcel(String sheet, int row, int cell) throws IOException{
    	 
    	 
    	 FileInputStream fis = new FileInputStream("C:\\Users\\AG22518\\Desktop\\bhavya\\input.xlsx");
         
         XSSFWorkbook wb = new XSSFWorkbook(fis);
         
         XSSFSheet sheet1 = wb.getSheet(sheet);
         
         XSSFRow row1 = sheet1.getRow(row);
          
         XSSFCell cell1 = row1.getCell(cell);
         
         String data = cell1.getStringCellValue();
 		         
         return data; 
    	 
    	 
     }
     public void enter_text(int row, int col, String x){
 		//Large_group_member c = new Large_group_member();
 		Property p1 = new Property(".startCol", col);
 		        Property p2 =  new Property(".startRow",row);   
 		        Property[] properties = {p1, p2};        
  		        sleep(0.2);
 		// TestObject[] lines = find(atChild(properties));        
 		TestObject[] lines = find(atDescendant(properties));
 		//sleep(1);
 		TextGuiTestObject text_box = ((TextGuiTestObject)lines[0]);
// 		text_box.waitForExistence(0.5,0.5);
 		//text_box.wait
 		text_box.waitForExistence(0.4,0.4);
 		text_box.setText(x);
// 		if(!x.equals("AG17470")){
// 		try {
// 			enter();
// 		} catch (AWTException e) {
// 			// TODO Auto-generated catch block
// 			e.printStackTrace();
// 		}
// 		}
// 		text_box.waitForExistence(0.4,0.4);
     }

     public void enter() throws AWTException{
 		Robot r = new Robot();
 		r.keyPress(KeyEvent.VK_ENTER);
 		r.keyRelease(KeyEvent.VK_ENTER);
 		
 		} 
     
     public void  F7() throws AWTException{
  		Robot r = new Robot();
  		r.keyPress(KeyEvent.VK_F7);
  		r.keyRelease(KeyEvent.VK_F7);
  		
  		}
     public void  F5() throws AWTException{
  		Robot r = new Robot();
  		r.keyPress(KeyEvent.VK_F5);
  		r.keyRelease(KeyEvent.VK_F5);
  		
  		}
  
     }
