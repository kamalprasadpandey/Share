package Pages;

import java.io.IOException;
import org.testng.annotations.AfterTest;
import org.testng.annotations.BeforeMethod;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.Test;
import java.awt.*;
import com.spire.xls.*;

import Pages.CommonFunctions;
public class ColourExcel {
	
	public static String projectPath=System.getProperty("user.dir");
	
	//https://www.e-iceblue.com/Tutorials/Java/Spire.XLS-for-Java/Program-Guide/Cells/Java-Set-Background-Color-and-Pattern-for-Excel-Cells.html
	@Test
	public void excel_version2() throws InterruptedException
	{
		 char c;
		 Workbook workbook = new Workbook();    
	           workbook.loadFromFile(projectPath+"/Data/Shares.xlsx");
	         //Get the first worksheet
	    	         Worksheet worksheet= workbook.getWorksheets().get(0);
	     //Set background color for range "A1:E1" and "A2:A10"
	    	      
	    	              worksheet.getRange().get("A3:R3").getStyle().setColor(Color.LIGHT_GRAY);
	    	              worksheet.getRange().get("A3:A20").getStyle().setColor(Color.yellow);
   	        for(c = 'B'; c < 'Q'; ++c)
   	        {
   	        	
   	        	for(int n=4;n<15;n++)
   	        	{
   	        	
   	        	  String mycell=c +""+ n;
   	        	  String mycell1=c +""+ (n+1);
   	        	//  System.out.println("mycell="+mycell);
   	        	// System.out.println("mycell1="+mycell1);
   	           
   	   
   	    	         // worksheet.getRange().get(mycell).getStyle().setColor(Color.PINK);
   	    	          
   	    	          
   	    	          //============================
   	    	       String value1=worksheet.getRange().get(mycell).getValue();
     	             float  value_1=Float.parseFloat(value1);  
     	             System.out.println("Value1="+value_1);
     	            
     	          String value2=worksheet.getRange().get(mycell1).getValue();
     	          float value_2=Float.parseFloat(value2);  
  	             System.out.println("Value2="+value_2);
     	           
  	             
     	                if (value_1>=value_2)
     	                {
     	                	
     	                 worksheet.getRange().get(mycell).getStyle().setColor(Color.GREEN);
    // System.out.println("green="+mycell);
     	  	             
     	                    }
     	                else
     	                {
     	                 worksheet.getRange().get(mycell1).getStyle().setColor(Color.RED);
     	                worksheet.getRange().get(mycell1).getStyle().setColor(Color.RED);
     	                System.out.println("Red="+mycell);
     	                }
     	               
     	  	             
   	    	          
   	        	}
   	        
   	        	
   	         workbook.saveToFile("CellBackground.xlsx", ExcelVersion.Version2013);
   	        }
   	  
	}
	
	
	
	
	
			
			public void excel_version1()
			{
				 char c;
				 Workbook workbook = new Workbook();    
	   	           workbook.loadFromFile(projectPath+"/Data/Shares.xlsx");
		   	        for(c = 'B'; c < 'N'; ++c)
		   	        {
		   	        	
		   	        	for(int n=3;n<15;n++)
		   	        	{
		   	        	
		   	        	  String mycell=c +""+ n;
		   	        	  String mycell1=c +""+ (n+1);
		   	        	  System.out.println(mycell);
		   	        	  
		   	           
		   	   //Get the first worksheet
		   	    	         Worksheet worksheet= workbook.getWorksheets().get(0);
		   	     //Set background color for range "A1:E1" and "A2:A10"
		   	    	      
		   	    	              worksheet.getRange().get("A3:N3").getStyle().setColor(Color.gray);
		   	    	              worksheet.getRange().get("A3:A15").getStyle().setColor(Color.yellow);
		   	    	          worksheet.getRange().get(mycell).getStyle().setColor(Color.GREEN);
		   	        	}
		   	        
		   	        	
		   	          //System.out.println(c + " ");
		   	        }
		   	     workbook.saveToFile("CellBackground.xlsx", ExcelVersion.Version2013);
			}
	
	public void ColourStocks() throws IOException
	{
			//Create a Workbook object
	        Workbook workbook = new Workbook();    
       //Load a sample Excel document
   	         workbook.loadFromFile(projectPath+"/Data/Shares.xlsx");
  //Get the first worksheet
   	         Worksheet worksheet= workbook.getWorksheets().get(0);
    //Set background color for range "A1:E1" and "A2:A10"
   	      
   	              worksheet.getRange().get("A3:N3").getStyle().setColor(Color.orange);
   	              worksheet.getRange().get("A3:A15").getStyle().setColor(Color.yellow);
   	      
   	     
   	         //---------------------------------------------------------------------------------
   	        	              
   	              
   	              
   	              String value1=worksheet.getRange().get("E7").getValue();
   	             float  value_1=Float.parseFloat(value1);  
   	             System.out.println("Value1="+value_1);
   	             
   	          String value2=worksheet.getRange().get("E8").getValue();
   	          float value_2=Float.parseFloat(value2);  
	             System.out.println("Value2="+value_2);
   	           
	             
   	                if (value_1>=value_2)
   	                {
   	                 worksheet.getRange().get("E7").getStyle().setColor(Color.GREEN);
   	                 System.out.println("Inside if");
   	                }
   	                else
   	                {
   	                 worksheet.getRange().get("E8").getStyle().setColor(Color.GREEN);
   	                }
   	             System.out.println("Inside else");
   	         //----------------------------------------------------------       
   	                  
   	                    //Save the document
   	           
   	                   workbook.saveToFile("CellBackground.xlsx", ExcelVersion.Version2013);

	}
	
	public void ColourStockscopy() throws IOException
	{
			//Create a Workbook object
	        Workbook workbook = new Workbook();    
       //Load a sample Excel document
   	         workbook.loadFromFile(projectPath+"/Data/Shares.xlsx");


   	      //Get the first worksheet
   	         Worksheet worksheet= workbook.getWorksheets().get(0);
   
   	         
   	    //Set background color for range "A1:E1" and "A2:A10"
   	      
   	              worksheet.getRange().get("A1:E1").getStyle().setColor(Color.green);
   	      
   	              worksheet.getRange().get("A2:A10").getStyle().setColor(Color.yellow);
   	      
   	       
   	         //Set background color for cell E8
   	           
   	                   worksheet.getRange().get("E8").getStyle().setColor(Color.red);
   	                worksheet.getRange().get("E9").getStyle().setColor(Color.GREEN);
   	           
   	                   //Set fill pattern style for range "C4:D5"
   	           
   	                   worksheet.getRange().get("C4:D5").getStyle().setFillPattern(ExcelPatternType.Percent25Gray);
   	           
   	                    //Save the document
   	           
   	                   workbook.saveToFile("CellBackground.xlsx", ExcelVersion.Version2013);
System.out.println("Completed");
	}
	

	
	
	@AfterTest
	public void after()
	{
		//driver.manage().deleteAllCookies();
		//driver.close();
		
	}

}
