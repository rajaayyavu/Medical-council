package org.hwru.selenium;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;
import java.util.concurrent.TimeUnit;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Row;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;

public class PopupData {

	public static void main(String[] args) throws InterruptedException, IOException, EncryptedDocumentException, InvalidFormatException {
		String[] links = null;
		int linksCount = 0;
		// TODO Auto-generated method stub
		System.setProperty("webdriver.chrome.driver", "C:\\Users\\preethalakshmi\\Downloads\\chromedriver.exe");

		WebDriver wb = new ChromeDriver();
		wb.manage().timeouts().implicitlyWait(20, TimeUnit.SECONDS);

		wb.manage().window().maximize();
		wb.get("https://www.mciindia.org/CMS/information-desk/indian-medical-register");
		Thread.sleep(5000);
		wb.findElement(By.linkText("Year of Registration")).click();
		//wb.findElement(By.id("doctor_year")).click();
		Thread.sleep(3000);
		wb.findElement(By.id("doctor_year")).sendKeys("2017");
		Thread.sleep(2000);
		wb.findElement(By.id("doctor_year_details")).click();
		Thread.sleep(3000);
		
//		//wb.get("http://www.mciindia.org/InformationDesk/IndianMedicalRegister.aspx");
//		//wb.findElement(By.xpath(".//*[@id='dnn_ctr588_IMRIndex_Link_Council']")).click();
//		wb.findElement(By.xpath(".//*[@id='dnn_ctr588_IMRIndex_Drp_StateCouncil']")).click();
//		wb.findElement(By.xpath(".//*[@id='dnn_ctr588_IMRIndex_Drp_StateCouncil']/option[35]")).click();
//		wb.findElement(By.xpath(".//*[@id='dnn_ctr588_IMRIndex_Submit_Btn']")).click();
//		wb.findElement(By.xpath(".//*[@id='dnn_ctr588_IMRIndex_GV_Search']/tbody/tr[1]/td/table/tbody/tr/td[1]/a")).click();
//		wb.findElement(By.xpath(".//*[@id='dnn_ctr588_IMRIndex_GV_Search']/tbody/tr[1]/td/table/tbody/tr/td[3]/a")).click();
//		wb.findElement(By.xpath(".//*[@id='dnn_ctr588_IMRIndex_GV_Search']/tbody/tr[1]/td/table/tbody/tr/td[3]/a")).click();
//		wb.findElement(By.xpath(".//*[@id='dnn_ctr588_IMRIndex_GV_Search']/tbody/tr[1]/td/table/tbody/tr/td[3]/a")).click();
//		wb.findElement(By.xpath(".//*[@id='dnn_ctr588_IMRIndex_GV_Search']/tbody/tr[1]/td/table/tbody/tr/td[3]/a")).click();
//		wb.findElement(By.xpath(".//*[@id='dnn_ctr588_IMRIndex_GV_Search']/tbody/tr[1]/td/table/tbody/tr/td[3]/a")).click();
//		wb.findElement(By.xpath(".//*[@id='dnn_ctr588_IMRIndex_GV_Search']/tbody/tr[1]/td/table/tbody/tr/td[3]/a")).click();
//		wb.findElement(By.xpath(".//*[@id='dnn_ctr588_IMRIndex_GV_Search']/tbody/tr[1]/td/table/tbody/tr/td[3]/a")).click();
//		wb.findElement(By.xpath(".//*[@id='dnn_ctr588_IMRIndex_GV_Search']/tbody/tr[1]/td/table/tbody/tr/td[3]/a")).click();
//		wb.findElement(By.xpath(".//*[@id='dnn_ctr588_IMRIndex_GV_Search']/tbody/tr[1]/td/table/tbody/tr/td[3]/a")).click();
//		wb.findElement(By.xpath(".//*[@id='dnn_ctr588_IMRIndex_GV_Search']/tbody/tr[1]/td/table/tbody/tr/td[3]/a")).click();
//		wb.findElement(By.xpath(".//*[@id='dnn_ctr588_IMRIndex_GV_Search']/tbody/tr[1]/td/table/tbody/tr/td[3]/a")).click();
//		wb.findElement(By.xpath(".//*[@id='dnn_ctr588_IMRIndex_GV_Search']/tbody/tr[1]/td/table/tbody/tr/td[3]/a")).click();
//		wb.findElement(By.xpath(".//*[@id='dnn_ctr588_IMRIndex_GV_Search']/tbody/tr[1]/td/table/tbody/tr/td[3]/a")).click();
//		wb.findElement(By.xpath(".//*[@id='dnn_ctr588_IMRIndex_GV_Search']/tbody/tr[1]/td/table/tbody/tr/td[3]/a")).click();
//		wb.findElement(By.xpath(".//*[@id='dnn_ctr588_IMRIndex_GV_Search']/tbody/tr[1]/td/table/tbody/tr/td[3]/a")).click();
//		wb.findElement(By.xpath(".//*[@id='dnn_ctr588_IMRIndex_GV_Search']/tbody/tr[1]/td/table/tbody/tr/td[3]/a")).click();
//		wb.findElement(By.xpath(".//*[@id='dnn_ctr588_IMRIndex_GV_Search']/tbody/tr[1]/td/table/tbody/tr/td[3]/a")).click();
//		wb.findElement(By.xpath(".//*[@id='dnn_ctr588_IMRIndex_GV_Search']/tbody/tr[1]/td/table/tbody/tr/td[3]/a")).click();
//		wb.findElement(By.xpath(".//*[@id='dnn_ctr588_IMRIndex_GV_Search']/tbody/tr[1]/td/table/tbody/tr/td[3]/a")).click();
//		wb.findElement(By.xpath(".//*[@id='dnn_ctr588_IMRIndex_GV_Search']/tbody/tr[1]/td/table/tbody/tr/td[3]/a")).click();
//		wb.findElement(By.xpath(".//*[@id='dnn_ctr588_IMRIndex_GV_Search']/tbody/tr[1]/td/table/tbody/tr/td[3]/a")).click();
//		wb.findElement(By.xpath(".//*[@id='dnn_ctr588_IMRIndex_GV_Search']/tbody/tr[1]/td/table/tbody/tr/td[3]/a")).click();
//		wb.findElement(By.xpath(".//*[@id='dnn_ctr588_IMRIndex_GV_Search']/tbody/tr[1]/td/table/tbody/tr/td[3]/a")).click();
//		wb.findElement(By.xpath(".//*[@id='dnn_ctr588_IMRIndex_GV_Search']/tbody/tr[1]/td/table/tbody/tr/td[3]/a")).click();
//		wb.findElement(By.xpath(".//*[@id='dnn_ctr588_IMRIndex_GV_Search']/tbody/tr[1]/td/table/tbody/tr/td[3]/a")).click();
//		wb.findElement(By.xpath(".//*[@id='dnn_ctr588_IMRIndex_GV_Search']/tbody/tr[1]/td/table/tbody/tr/td[3]/a")).click();
//		wb.findElement(By.xpath(".//*[@id='dnn_ctr588_IMRIndex_GV_Search']/tbody/tr[1]/td/table/tbody/tr/td[3]/a")).click();
//		wb.findElement(By.xpath(".//*[@id='dnn_ctr588_IMRIndex_GV_Search']/tbody/tr[1]/td/table/tbody/tr/td[3]/a")).click();
//		wb.findElement(By.xpath(".//*[@id='dnn_ctr588_IMRIndex_GV_Search']/tbody/tr[1]/td/table/tbody/tr/td[3]/a")).click();
//		wb.findElement(By.xpath(".//*[@id='dnn_ctr588_IMRIndex_GV_Search']/tbody/tr[1]/td/table/tbody/tr/td[3]/a")).click();
//		wb.findElement(By.xpath(".//*[@id='dnn_ctr588_IMRIndex_GV_Search']/tbody/tr[1]/td/table/tbody/tr/td[3]/a")).click();
//		wb.findElement(By.xpath(".//*[@id='dnn_ctr588_IMRIndex_GV_Search']/tbody/tr[1]/td/table/tbody/tr/td[3]/a")).click();
//		wb.findElement(By.xpath(".//*[@id='dnn_ctr588_IMRIndex_GV_Search']/tbody/tr[1]/td/table/tbody/tr/td[3]/a")).click();
//		wb.findElement(By.xpath(".//*[@id='dnn_ctr588_IMRIndex_GV_Search']/tbody/tr[1]/td/table/tbody/tr/td[3]/a")).click();
//		wb.findElement(By.xpath(".//*[@id='dnn_ctr588_IMRIndex_GV_Search']/tbody/tr[1]/td/table/tbody/tr/td[3]/a")).click();
//		wb.findElement(By.xpath(".//*[@id='dnn_ctr588_IMRIndex_GV_Search']/tbody/tr[1]/td/table/tbody/tr/td[3]/a")).click();
//		wb.findElement(By.xpath(".//*[@id='dnn_ctr588_IMRIndex_GV_Search']/tbody/tr[1]/td/table/tbody/tr/td[3]/a")).click();
//		wb.findElement(By.xpath(".//*[@id='dnn_ctr588_IMRIndex_GV_Search']/tbody/tr[1]/td/table/tbody/tr/td[3]/a")).click();
//		wb.findElement(By.xpath(".//*[@id='dnn_ctr588_IMRIndex_GV_Search']/tbody/tr[1]/td/table/tbody/tr/td[3]/a")).click();
//		wb.findElement(By.xpath(".//*[@id='dnn_ctr588_IMRIndex_GV_Search']/tbody/tr[1]/td/table/tbody/tr/td[3]/a")).click();
//		wb.findElement(By.xpath(".//*[@id='dnn_ctr588_IMRIndex_GV_Search']/tbody/tr[1]/td/table/tbody/tr/td[3]/a")).click();
//		wb.findElement(By.xpath(".//*[@id='dnn_ctr588_IMRIndex_GV_Search']/tbody/tr[1]/td/table/tbody/tr/td[3]/a")).click();
//		wb.findElement(By.xpath(".//*[@id='dnn_ctr588_IMRIndex_GV_Search']/tbody/tr[1]/td/table/tbody/tr/td[3]/a")).click();
//		wb.findElement(By.xpath(".//*[@id='dnn_ctr588_IMRIndex_GV_Search']/tbody/tr[1]/td/table/tbody/tr/td[3]/a")).click();
//		wb.findElement(By.xpath(".//*[@id='dnn_ctr588_IMRIndex_GV_Search']/tbody/tr[1]/td/table/tbody/tr/td[3]/a")).click();
//		wb.findElement(By.xpath(".//*[@id='dnn_ctr588_IMRIndex_GV_Search']/tbody/tr[1]/td/table/tbody/tr/td[3]/a")).click();
//		wb.findElement(By.xpath(".//*[@id='dnn_ctr588_IMRIndex_GV_Search']/tbody/tr[1]/td/table/tbody/tr/td[3]/a")).click();
//		wb.findElement(By.xpath(".//*[@id='dnn_ctr588_IMRIndex_GV_Search']/tbody/tr[1]/td/table/tbody/tr/td[3]/a")).click();
//		wb.findElement(By.xpath(".//*[@id='dnn_ctr588_IMRIndex_GV_Search']/tbody/tr[1]/td/table/tbody/tr/td[3]/a")).click();
//		wb.findElement(By.xpath(".//*[@id='dnn_ctr588_IMRIndex_GV_Search']/tbody/tr[1]/td/table/tbody/tr/td[3]/a")).click();
//		wb.findElement(By.xpath(".//*[@id='dnn_ctr588_IMRIndex_GV_Search']/tbody/tr[1]/td/table/tbody/tr/td[3]/a")).click();
//		wb.findElement(By.xpath(".//*[@id='dnn_ctr588_IMRIndex_GV_Search']/tbody/tr[1]/td/table/tbody/tr/td[3]/a")).click();

		/*for(int i1=0;i1<=50;i1++)
		{*/
		WebElement elm=wb.findElement(By.xpath(".//*[@id='dnn_ctr588_IMRIndex_GV_Search']"));
		System.out.println(elm.getText());
		//Thread.sleep(3000);
		
		
		List<WebElement> alllinks = elm.findElements(By.tagName("a")); 
		linksCount = alllinks.size();
		System.out.println("Total no of links Available: "+linksCount);
	
		
		 String a[]=new String[alllinks.size()];
		 Thread.sleep(3000);
		 int rowCount = 26664;
		  for(int i=167;i<alllinks.size();i++)
	        {
	           if(!(alllinks.get(i).getText().isEmpty()))
	            {
	        	   
	                        	            		
	            		System.out.println(alllinks.get(i).getText());
	            		
	            		a[i]=alllinks.get(i).getText(); 
	            		
	            		 if(a[i].startsWith("V"))
	                     {
	                         //System.out.println("clicking on this link::"+driver.findElement(By.linkText(a[i])).getText());
	            			 Thread.sleep(2000);
	            			 alllinks.get(i).click(); 
	                         
	                         String winHandleBefore = wb.getWindowHandle();

	     			 		// Perform the click operation that opens new window

	     			 		// Switch to new window opened
	     			 		for(String winHandle : wb.getWindowHandles()){
	     			 			wb.switchTo().window(winHandle);
	     			 			
	     			 		}
	     			 		
	     			 		//Thread.sleep(3000);
	                         WebElement elm1=wb.findElement(By.xpath(".//*[@id='Name']"));
	                         WebElement elm2=wb.findElement(By.xpath(".//*[@id='DOB']"));
	                         WebElement elm3=wb.findElement(By.xpath(".//*[@id='FatherName']"));
	                         WebElement elm4=wb.findElement(By.xpath(".//*[@id='Address']"));
	                         WebElement elm5=wb.findElement(By.xpath(".//*[@id='Lbl_Council']"));
	                         WebElement elm6=wb.findElement(By.xpath(".//*[@id='Regis_no']"));
	                         WebElement elm7=wb.findElement(By.xpath(".//*[@id='Qual']"));
	                         WebElement elm8=wb.findElement(By.xpath(".//*[@id='QualYear']"));
	                         WebElement elm9=wb.findElement(By.xpath(".//*[@id='Univ']"));
		    			 		System.out.println(elm1.getText()+" "+elm2.getText()+" "+elm3.getText()+" "+elm4.getText()+" "+elm5.getText()+" "+elm6.getText()+" "+elm7.getText()+" "+elm8.getText()+" "+elm9.getText());
		    			 		  
		    			 		 try {
		      						  // Specify the path of file
		    			 			   File src=new File("C:\\Users\\preethalakshmi\\Downloads\\selenium1.xls");
		      						 
		      						   // load file
		      						   FileInputStream fis=new FileInputStream(src);
		      						 
		      						   // Load workbook
		      						   HSSFWorkbook wb1=new HSSFWorkbook(fis);
		      						   
		      						   // Load sheet- Here we are loading first sheetonly
		      						      HSSFSheet sh1= wb1.getSheetAt(0);
		      						    
				    					/*Row row1 = sh1.createRow(0);
				    					row1.createCell(0).setCellValue(elm1.getText());
				    					row1.createCell(1).setCellValue(elm2.getText());*/
				    					
				    					
		      						 
		      						  // getRow() specify which row we want to read.
		      						 
		      						  // and getCell() specify which column to read.
		      						  // getStringCellValue() specify that we are reading String data.
		      						    
		      						      //sh1.getRow(0).createCell(2).setCellValue(elm1.getText());
		      						     // System.out.println(sh1.getRow(0).getCell(0).getStringCellValue());
		      						     /* System.out.println(sh1.getRow(0).getCell(0).getStringCellValue());
		      						      
		      						      System.out.println(sh1.getRow(0).getCell(1).getStringCellValue());
		      						      
		      						      System.out.println(sh1.getRow(1).getCell(0).getStringCellValue());
		      						      
		      						      System.out.println(sh1.getRow(1).getCell(1).getStringCellValue());
		      						      
		      						      System.out.println(sh1.getRow(2).getCell(0).getStringCellValue());
		      						      
		      						      System.out.println(sh1.getRow(2).getCell(1).getStringCellValue());
		      						      
		      						 
		      						
		      						      
		      						 sh1.getRow(0).createCell(2).setCellValue(elm1.getText());
		      						 
		      						 sh1.getRow(0).createCell(3).setCellValue(elm2.getText());
		      						 
		      						 sh1.getRow(0).createCell(4).setCellValue(elm3.getText());
		      						 sh1.getRow(0).createCell(5).setCellValue(elm4.getText());
		      						 sh1.getRow(0).createCell(6).setCellValue(elm5.getText());
		      						 sh1.getRow(0).createCell(7).setCellValue(elm6.getText());
		      						 sh1.getRow(0).createCell(8).setCellValue(elm7.getText());
		      						 sh1.getRow(0).createCell(9).setCellValue(elm8.getText());
		      						 sh1.getRow(0).createCell(10).setCellValue(elm9.getText());*/
		                               
		      						 
		      						
		      						 
		      						 
		      						// here we need to specify where you want to save file
		      						    
		      						//Iterator<Row> rows=sh1.rowIterator();
		      						 /*while(rows.hasNext()){*/
		      						     
		      						 
		      						     
		      						  Row row11 = sh1.createRow(rowCount);
		      						    	//System.out.println(sh1.getRow(0).getCell(0).getStringCellValue());
		      						      
		     
		      						      
		      							// Row currentRow=rows.next();
		      							// System.out.println(currentRow.getCell(0).getStringCellValue());
		      							 //System.out.println(currentRow.getCell(1).getStringCellValue());
		      							
		      						 
		      							 
		      							 for(int cellPlace=0;cellPlace<=8;cellPlace++){
		      								 switch(cellPlace){
		      								 case 0:
		      									row11.createCell(cellPlace).setCellValue(elm1.getText());
		      									 break;
		      								 case 1:
		      									row11.createCell(cellPlace).setCellValue(elm2.getText());
		      									 break;
		      								 case 2:
		      									row11.createCell(cellPlace).setCellValue(elm3.getText());
		      									 break;
		      								 case 3:
		      									row11.createCell(cellPlace).setCellValue(elm4.getText());
		      									 break;
		      								 case 4:
		      									row11.createCell(cellPlace).setCellValue(elm5.getText());
		      									 break;
		      								 case 5:
		      									row11.createCell(cellPlace).setCellValue(elm6.getText());
		      									 break;
		      								 case 6:
		      									row11.createCell(cellPlace).setCellValue(elm7.getText());
		      									 break;
		      								 case 7:
		      									row11.createCell(cellPlace).setCellValue(elm8.getText());
		      									 break;
		      								 case 8:
		      									row11.createCell(cellPlace).setCellValue(elm9.getText());
		      									 break;
		      								 }
		      							 }
		      							 
		      							
			      						 
		      						/* }*/
		      						  rowCount++;
                                      
                                      FileOutputStream fout=new FileOutputStream(new File("C:\\Users\\preethalakshmi\\Downloads\\selenium1.xls"));
                                      
                                      
                                     // finally write content
                                      
                                      wb1.write(fout);
                                     fout.close();
		      						 
		      						 
		      						 
		      							 
		      							 
		      						// close the file
		      						 
		      						 
		      						 
		      						 } catch (Exception e) {
		      						 
		      						   System.out.println(e.getMessage());
		      						 
		      						  }
		    			 		
		    					

		    			 
		    			 		
		    	    			wb.close();
		    	           		 wb.switchTo().window(winHandleBefore);
	                     }
	            		/* else{
	            			 a[i].startsWith("N");
	            			 alllinks.get(i).click();
	            		 }*/
	            
	            		
   			 		 
	            		 
	            		 
	            			
	            		
	            
	            	
	            }     
	       
	
	
		
	        }
		  /*wb.findElement(By.xpath(".//*[@id='dnn_ctr588_IMRIndex_GV_Search']/tbody/tr[1]/td/table/tbody/tr/td[3]/a")).click();
		 // wb.findElement(By.xpath(".//*[@id='dnn_ctr588_IMRIndex_GV_Search']/tbody/tr[503]/td/table/tbody/tr/td[3]/a")).click();
		}*/
	}
	
}
