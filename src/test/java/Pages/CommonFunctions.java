package Pages;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.concurrent.TimeUnit;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.support.ui.Select;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
public class CommonFunctions {

	public static WebDriver driver=null;
	public static String projectPath=System.getProperty("user.dir");
	
	private String baseUrl;
	public CommonFunctions(WebDriver driver)
	{
		this.driver=driver;
	}
	

	public void highLighterMethod(WebDriver driver, WebElement element){
		JavascriptExecutor js = (JavascriptExecutor) driver;
		js.executeScript("arguments[0].setAttribute('style', 'background: yellow; border: 2px solid red;');", element);
		}
	
	public static void Search_Historical_Price(String Equity,int execol) throws InterruptedException {
		System.setProperty("webdriver.chrome.driver", projectPath+"/drivers/chromedriver/chromedriver.exe");   // Code to launch chrome driver
		//-------- Chrome options for handling Exception alert Pop-up 
		ChromeOptions options = new ChromeOptions();
		options.setExperimentalOption("useAutomationExtension", false);
		driver = new ChromeDriver(options);
		driver.manage().timeouts().implicitlyWait(7,TimeUnit.SECONDS) ;
		//------- End of Chrome options  -----------		
		driver.navigate().to("https://www.moneycontrol.com/stocks/histstock.php?");   
		driver.manage().window().maximize();
		driver.manage().timeouts().implicitlyWait(15,TimeUnit.SECONDS) ;	
		Thread.sleep(5000);
		
		driver.findElement(By.id("wzrk-cancel")).click();
		Thread.sleep(2000);
		driver.findElement(By.xpath("//input[@class='txtsrchbox FL']")).click();
		Thread.sleep(2000);
		driver.navigate().refresh();
		Thread.sleep(5000);
		if(execol<=3)
		{
			driver.findElement(By.id("wutabs2")).click();
			driver.manage().timeouts().implicitlyWait(10,TimeUnit.SECONDS) ;
			
			
			Select from_index = new Select(driver.findElement(By.name("indian_indices")));
			from_index.selectByVisibleText(Equity);//S&P BSE Sensex
		}
		else
				{
				driver.findElement(By.id("mycomp")).sendKeys(Equity);
				driver.findElement(By.id("mycomp")).click();
				Thread.sleep(4000);	
				WebElement element = driver.findElement(By.xpath("//div[@id='suggest']/ul/li[1]"));
				
			   // String text = element.getText();
			   	driver.findElement(By.xpath("//div[@id='suggest']/ul/li[1]")).click();
				Select exchange = new Select(driver.findElement(By.name("ex")));
				exchange.selectByVisibleText("NSE");
				}
		
		Select from_day = new Select(driver.findElement(By.name("frm_dy")));
		from_day.selectByVisibleText("20");
		
		Select from_month = new Select(driver.findElement(By.name("frm_mth")));
		from_month.selectByVisibleText("Mar");
		
		Select from_year = new Select(driver.findElement(By.name("frm_yr")));
		from_year.selectByVisibleText("2023");
		
		Select to_day = new Select(driver.findElement(By.name("to_dy")));
		to_day.selectByVisibleText("06");
		
		Select to_month = new Select(driver.findElement(By.name("to_mth")));
		to_month.selectByVisibleText("Apr");
		
		Select to_year = new Select(driver.findElement(By.name("to_yr")));
		to_year.selectByVisibleText("2023");
		
		driver.findElement(By.xpath("//tbody/tr[1]/td[1]/form[1]/div[4]/input[1]")).click();
		driver.manage().timeouts().implicitlyWait(20,TimeUnit.SECONDS) ;
		driver.manage().timeouts().implicitlyWait(20,TimeUnit.SECONDS) ;
		
		copyTable(Equity,execol,driver);		// call function to copy equity data 
		System.out.println(Equity+" "+(execol-1)+" Completed");
		
		driver.manage().timeouts().implicitlyWait(20,TimeUnit.SECONDS) ;
		driver.close(); //Close browser
	
	}
	
	
	public static void copyTable(String Equity,int excelcol,WebDriver driver) {
		int rowheader=2;
		int date_col=0;
		//No of cols	 
		java.util.List<WebElement>  col = driver.findElements(By.xpath("/html[1]/body[1]/div[6]/div[2]/div[1]/div[4]/div[4]/table[1]/tbody[1]/tr/th"));
		System.out.println("No of cols are : " +col.size()); 
        //No.of rows 
		java.util.List<WebElement>  rows = driver.findElements(By.xpath("/html[1]/body[1]/div[6]/div[2]/div[1]/div[4]/div[4]/table[1]/tbody[1]/tr")); 
        System.out.println("No of rows are : " + rows.size());
        	        
         //-----------------------------
       
		int rowlen=rows.size()+1;
		
		
        for (int i=3;i<rowlen;i++)
        	{
        	
		    WebElement cellIneed = driver.findElement(By.xpath("/html[1]/body[1]/div[6]/div[2]/div[1]/div[4]/div[4]/table[1]/tbody[1]/tr["+i+"]/td[5]"));
		    String valueIneed = cellIneed.getText();
		    System.out.println("Cell value is : " + valueIneed); 
		    //--------Code to write data in excel
		 		
		    try{
				XSSFWorkbook wb= null;
				XSSFSheet sh=null;
				
				File file= new File(projectPath+"/Data/Shares.xlsx");
				//File file= new File("F:\\Shares.xlsx");
				
				FileInputStream fis = new FileInputStream(file);
				wb= new XSSFWorkbook(fis);
				sh=wb.getSheetAt(0);
		 
				//----<<-- code to write date in excel
				
				if(date_col==0)
				{
					
				sh.getRow(2).createCell(date_col).setCellValue("Date");
				int date_row;
				for(date_row=3;date_row<rowlen; date_row++)
				{
					WebElement dateneed = driver.findElement(By.xpath("/html[1]/body[1]/div[6]/div[2]/div[1]/div[4]/div[4]/table[1]/tbody[1]/tr["+date_row+"]/td[1]"));
				    String DateIneed = dateneed.getText();
				    /*
				    System.out.println(" Date is : " + DateIneed); 
				    
				   
				    System.out.println("date_row= "+ date_row);
				    System.out.println("date_col= "+ date_col);
				    System.out.println("DateIneed= "+ DateIneed);
				  */
					sh.getRow(date_row).createCell(date_col).setCellValue(DateIneed);
					//System.out.println(" Date is written in Excel"); 
				}
				
				date_col=2; // It will increase row header count which will be written only once
				}
		//
       
				
			
				//----code to write date in excel ----->>
				//<------ code to write Equity name in excel as header
				if(rowheader==2)
				{
				sh.getRow(rowheader).createCell(excelcol).setCellValue(Equity);	
				//System.out.println(" Header is written in Excel"); 
				rowheader=3; // It will increase row header count which will be written only once
				}
				//------ code to write Equity name in excel as header	--------->>			
				
				//sh.getRow(2).createCell(1).setCellValue(valueIneed);
		sh.getRow(i).createCell(excelcol).setCellValue(valueIneed);
		//System.out.println(Equity+" "+excelcol+" Completed");
		//System.out.println(valueIneed);
		//System.out.println(" Value is written in Excel"); 
		FileOutputStream fso =new FileOutputStream(file);
		wb.write(fso);
		
		
	}
		catch(Exception e1)	{
			e1.printStackTrace();
		}
		    
		   // System.out.println("Completed");
        	}
		
	}
		

	


	public void myWait(int myTime)
	{
		
		try {
            
                Thread.sleep(myTime);
        }
        catch (InterruptedException e) {
            System.out.println("thread interrupted");
        }
    
	}
	
	public void colourexcel()
	{
		//https://www.e-iceblue.com/Tutorials/Java/Spire.XLS-for-Java/Program-Guide/Cells/Java-Set-Background-Color-and-Pattern-for-Excel-Cells.html
		
		
		
		
	}
	}

