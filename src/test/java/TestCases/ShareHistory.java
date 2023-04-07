package TestCases;

import java.io.File;
import java.io.IOException;
import java.lang.reflect.Method;
import java.util.concurrent.TimeUnit;
import org.apache.commons.io.FileUtils;
import org.openqa.selenium.*;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.support.ui.Select;
import org.testng.annotations.AfterTest;
import org.testng.annotations.BeforeMethod;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.Test;
import com.relevantcodes.extentreports.ExtentReports;
import com.relevantcodes.extentreports.ExtentTest;
import com.relevantcodes.extentreports.LogStatus;
/*
//---- for colour
//https://www.geeksforgeeks.org/how-to-fill-background-color-of-cells-in-excel-using-java-and-apache-poi/
import org.testng.annotations.Test;
import java.io.FileOutputStream;
import java.io.IOException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
*/

import Pages.CommonFunctions;
public class ShareHistory {
	

	public WebDriver driver;
	static ExtentReports report;
	static ExtentTest test;
	
	@BeforeTest
	public void setup () throws Exception
	{
		
	String projectPath=System.getProperty("user.dir");
	/*//chrome setup
	System.setProperty("webdriver.chrome.driver",projectPath+"/drivers/chromedriver/chromedriver.exe");
	ChromeOptions options=new ChromeOptions();
	driver =new ChromeDriver(options);
	*/
	report= new ExtentReports(projectPath+"/ExtentReport/report.html",true);
	test= report.startTest("Extent Reports");
	}
	
	@BeforeMethod
	public void beformethod(Method method) throws Exception
	{
		test=report.startTest((this.getClass().getSimpleName()+" :: " + method.getName()), method.getName());
		test.assignAuthor("Kamal Pandey");
		test.assignCategory("Shares-History");
		test.log(LogStatus.PASS,"Browser Launched Successfully");
		test.log(LogStatus.PASS, method.getName()+"  Execution Started  ");
		
		}
	

	@Test
	public void TC_Index() throws Exception
	{
		CommonFunctions cf= new CommonFunctions(driver);
		test.log(LogStatus.INFO, "Test Case Validation Started:Launch Application");
		int Dashboard_j =10;
		int copydate=1;
		String[] stockname= {"S&P BSE Sensex","CNX Nifty","Bank Nifty"};
				 
		//------------------Update value of i as per array/stocks index , it begins from 0---------------------
		for (int i=0; i<stockname.length; i++) 
		{ 
		cf.Search_Historical_Price(stockname[i],i+1);
			}	
		//test.log(LogStatus.INFO,test.addScreenCapture(CaptureScreen(driver))+"Application Launched Successfully" );
		test.log(LogStatus.PASS, "1st Test Case Executed Successfully");
	}
	@Test
	public void TC_FMCG() throws Exception
	{
		CommonFunctions cf= new CommonFunctions(driver);
		test.log(LogStatus.INFO, "Test Case Validation Started:Launch Application");
		int Dashboard_j =10;
		int copydate=1;
		//String[] stockname= {"Dabur India"};
		String[] stockname= {"Dabur India","Britannia Industries","Hindustan Unilever","TATA Consumer Products"};
			//------------------Update value of i as per array/stocks index , it begins from 0---------------------
			 
		//------------------Update value of i as per array/stocks index , it begins from 0---------------------
		for (int i=0; i<stockname.length; i++) 
		{ 
		cf.Search_Historical_Price(stockname[i],i+4);
			}	
		
		
		//test.log(LogStatus.INFO,test.addScreenCapture(CaptureScreen(driver))+"Application Launched Successfully" );
		test.log(LogStatus.PASS, "1st Test Case Executed Successfully");
	}
	@Test
	public void TC_MustHave() throws Exception
	{
		CommonFunctions cf= new CommonFunctions(driver);
		test.log(LogStatus.INFO, "Test Case Validation Started:Launch Application");
		int Dashboard_j =10;
		int copydate=1;
		
		String[] stockname= {"Tata Motors","Asian Paints","Titan Company","Reliance Industries","Tata Consultancy Services"};

		//------------------Update value of i as per array/stocks index , it begins from 0---------------------
			 
		//------------------Update value of i as per array/stocks index , it begins from 0---------------------
		for (int i=0; i<stockname.length; i++) 
		{ 
		cf.Search_Historical_Price(stockname[i],i+8);
			}	
		
	}
	@Test
	public void Banks() throws Exception
	{
		CommonFunctions cf= new CommonFunctions(driver);
		test.log(LogStatus.INFO, "Test Case Validation Started:Launch Application");
		int Dashboard_j =10;
		int copydate=1;
		
		String[] stockname= {"ICICI Bank","HDFC Bank","State Bank Of India"};

		//------------------Update value of i as per array/stocks index , it begins from 0---------------------
			 
		//------------------Update value of i as per array/stocks index , it begins from 0---------------------
		for (int i=0; i<stockname.length; i++) 
		{ 
		cf.Search_Historical_Price(stockname[i],i+13);
			}	
		
	}
	
	public static String CaptureScreen(WebDriver driver) throws IOException
	{
		File srcfile = ((TakesScreenshot)driver).getScreenshotAs(OutputType.FILE);
		 String path=System.getProperty("user.dir")+"\\Screenshots\\Screenshot_"+System.currentTimeMillis()+".png";
		 File Destinationfile=new File(path);
		
		String absolutepath_screen = Destinationfile.getAbsolutePath();
		FileUtils.copyFile(srcfile,Destinationfile);
		return absolutepath_screen;
	}
	
	
	@AfterTest
	public void after()
	{
		/*
		//driver.manage().deleteAllCookies();
		driver.close();
		report.endTest(test);
		report.flush();
		*/
	}

}
