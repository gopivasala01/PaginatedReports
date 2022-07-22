package Package1;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.time.Duration;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.concurrent.TimeUnit;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.Keys;
import org.openqa.selenium.Proxy;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.edge.EdgeDriver;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.remote.CapabilityType;
import org.openqa.selenium.remote.DesiredCapabilities;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;

import Package1.AppConfig;
import Package1.RunnerClass;

public class RunnerClass {

	static	ChromeDriver driver;
	static	EdgeDriver driver_Edge;
	static Actions actions;
	static JavascriptExecutor js;
	static File file;
	static FileInputStream fis;
	static StringBuilder stringBuilder = new StringBuilder() ;
	static WebDriverWait wait;
	static int flag =0;
	static File userCredentialsFile;
	static FileInputStream userFileInputStream;
	static Calendar userLoggedInTime;
	static FileOutputStream fos;
	static int lastRowNumber;
	static boolean flagToCheckUserLoginTime = false;
	static String loggedInTimeInSheet;
	static boolean flagToCheckLoginForFirstTimeOrRepeatedUser;
	static File agentEmails;
	static FileInputStream agentEmailFileInputStream;
	static FileOutputStream agentEmailFileOutputStream;
	static ArrayList<String> emailsList;
	static RunnerClass runnerClassObject;
	static SendAttachmentInMail mailObject;
	RunnerClass()
	{
		System.setProperty(AppConfig.browserType,AppConfig.browserPath);
		Map<String, Object> prefs = new HashMap<String, Object>();
        
        // Use File.separator as it will work on any OS
        prefs.put("download.default_directory",
                System.getProperty("user.dir") + File.separator + "externalFiles" + File.separator + "downloadFiles");
         
        // Adding cpabilities to ChromeOptions
        ChromeOptions options = new ChromeOptions();
        options.setExperimentalOption("prefs", prefs);
         
        // Printing set download directory
         
        // Launching browser with desired capabilities
         driver= new ChromeDriver(options);
		//driver_Edge = new EdgeDriver();
		js = (JavascriptExecutor)driver;
		actions = new Actions(driver);
		//js = (JavascriptExecutor)driver_Edge;
		//actions = new Actions(driver_Edge);
		//wait = new WebDriverWait(driver,100);
		driver.manage().timeouts().implicitlyWait(15,TimeUnit.SECONDS);
		wait = new WebDriverWait(driver, Duration.ofSeconds(15));
		emailsList = new ArrayList<String>();
	}
	public static void main(String[] args) throws Exception
	{
		GetDataFromDatabase databaseClassObject =new  GetDataFromDatabase();
	    //databaseClassObject.getData();
		//System.out.println("Data exported from Database");
		
		runnerClassObject = new RunnerClass();
		mailObject = new SendAttachmentInMail();
		//OutlookMailForward outlookMailObject = new OutlookMailForward();
		//OutlookMailForward.loginToMail();
	   driver.get(AppConfig.powerBIURL);
	    driver.manage().window().maximize();
	    runnerClassObject.login();
	    //runnerClassObject.navigateToReport();
	    runnerClassObject.readDataFromExcel();
	    Thread.sleep(8000);
	    //runnerClassObject.getReportExcel();
	    //mailObject.sendFileToMail();
	    
		 
	    
	    
	}
	public void downloadEachReportOfAgent()
	{
		driver.navigate().to(AppConfig.reportLink);
	}
	public void login() throws Exception
	{
		try {
		driver.findElement(Locators.userName).sendKeys(AppConfig.userName);
		driver.findElement(Locators.loginButton).click();
		}catch(Exception e) {
			driver.findElement(Locators.userName2).sendKeys(AppConfig.userName);
			driver.findElement(Locators.loginButton2).click();
		}
		Thread.sleep(2000);
		driver.findElement(Locators.password).sendKeys(AppConfig.password);
		driver.findElement(Locators.signInButton).click();
		actions.sendKeys(Keys.ENTER).build().perform();
		Thread.sleep(5000);
		
	}
	public void navigateToReport() throws Exception
	{
		Thread.sleep(3000);
		driver.findElement(Locators.workspacesButton).click();
		List<WebElement> workspacesList = driver.findElements(Locators.workspacesList);
		/*for(int i=0;i<workspacesList.size();i++)
		{
			String text = workspacesList.get(i).getText();
			if(text.equals(AppConfig.workspaceName));
			{
				workspacesList.get(i).click();
				break;
			}
		}*/
		driver.findElement(By.xpath("(//*[@class='workspaceName'])[last()]")).click();
		Thread.sleep(5000);
		List<WebElement> reportList = driver.findElements(Locators.reportList);
		for(int i=0;i<reportList.size();i++)
		{
			String text = reportList.get(i).getText();
			if(text.equalsIgnoreCase(AppConfig.reportName));
			{
				reportList.get(i).click();
				break;
			}
		}
		Thread.sleep(5000);
		driver.switchTo().frame(driver.findElement(Locators.reportFrame));
		driver.findElement(Locators.leasingAgentDropdown).click();
		Thread.sleep(2000);
		List<WebElement> leasingAgentNames = driver.findElements(Locators.leasingAgentsList);
		List<WebElement> leasingAgentCheckboxes = driver.findElements(Locators.leasingAgentsListCheckboxes);
		System.out.println(leasingAgentNames.size());
		//driver.findElement(Locators.leasingAgentsList).click();
		for(int i=0;i<leasingAgentNames.size();i++)
		{
			String text = leasingAgentNames.get(i).getText();
			if(text.contains("Buckley"))
			{
				actions.moveToElement(leasingAgentCheckboxes.get(i)).click(leasingAgentCheckboxes.get(i)).build().perform();
				break;
			}
			if(i==leasingAgentNames.size()-1)
			{
				
				//actions.moveToElement(leasingAgentCheckboxes.get(i)).build().perform();
				leasingAgentNames = driver.findElements(Locators.leasingAgentsList);
				leasingAgentCheckboxes = driver.findElements(Locators.leasingAgentsListCheckboxes);
				System.out.println(leasingAgentNames.size());
				i=0;
			}
		}
		driver.findElement(Locators.clickOutsideOfTheReport).click();
		driver.findElement(Locators.viewReportButton).click();
		Thread.sleep(5000);
		driver.findElement(Locators.exportButton).click();
		driver.findElement(Locators.excelButton).click();
	
		
	
	}
	public void getReportExcel(String agentName) throws Exception
	{
		//driver.navigate().to(AppConfig.reportLink+"Aaron Sonnier");
		//driver.switchTo().frame(driver.findElement(Locators.reportFrame));
		//driver.findElement(Locators.clickOutsideOfTheReport).click();
		//driver.findElement(Locators.viewReportButton).click();
		Thread.sleep(5000);
		wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("(//*[@aria-label='Report Text'])[1]")));
		driver.findElement(Locators.exportButton).click();
		Thread.sleep(2000);
		driver.findElement(Locators.excelButton).click();
		Thread.sleep(10000);
		// Rename the file name
		/*
		File downloadPath = new File (AppConfig.downloadPath);
		Thread.sleep(2000);
		//File renamedFile = new File ("C:\\Gopi\\Automation\\Eclipse Workspace\\PowerBIReportSubscriptions\\externalFiles\\downloadFiles\\Marketing by LeasingAgent_"+agentName+".xlsx");
		downloadPath.renameTo(new File (downloadPath+"_"+agentName+".xlsx"));
		System.out.println("File path with name"+downloadPath);
		//downloadPath.delete();
		*/
		File file = new File (AppConfig.downloadPath);
		File  File2 = new File("C:\\Gopi\\Projects\\Property ware\\Paginated Reports\\"+"_"+agentName+".xlsx");
		System.out.println(file.exists());
		
		java.io.FileWriter out= new java.io.FileWriter(File2, true /*append=yes*/);
		
		if(file.exists())
		{
			file.delete();
		}
		
	}
	
	public void readDataFromExcel() throws Exception
	{
		agentEmails= new File(AppConfig.userEmailsFilePath);
		agentEmailFileInputStream = new FileInputStream(agentEmails);
		
		XSSFWorkbook workbook = new XSSFWorkbook(agentEmailFileInputStream);
		XSSFSheet sheet = workbook.getSheetAt(0);
		for(int i=0;i<sheet.getLastRowNum();i++)
		{
			//for specific Companies
			Row row_Company = sheet.getRow(i);
			Cell cell_Company = row_Company.getCell(2);
			if(cell_Company.toString().equalsIgnoreCase("OKC")||cell_Company.toString().equalsIgnoreCase("Tulsa"))
			{
			
			Row row_Status = sheet.getRow(i);
			Cell cell_Status = row_Status.getCell(5);
			if(cell_Status.toString().equalsIgnoreCase("Pending"))
			{
			Row row_Name = sheet.getRow(i);
			Cell cell_Name = row_Name.getCell(0);
			String agentName = cell_Name.toString();
			System.out.println(agentName);
			driver.navigate().to(AppConfig.reportLink+agentName);
			try
			{
				wait.until(ExpectedConditions.frameToBeAvailableAndSwitchToIt(Locators.reportFrame));
				wait.until(ExpectedConditions.invisibilityOfElementLocated(Locators.spinner));
				try
				{
				if(driver.findElement(Locators.reportExists).isDisplayed())
				{
				runnerClassObject.getReportExcel(agentName);
				Thread.sleep(10000);
				Row row_Email = sheet.getRow(i);
				
				//Agent Name and agent Email
				Cell cell_Email = row_Email.getCell(1);
				String emailText = cell_Email.toString();
				System.out.println(emailText);
				String agentEmail=null ;
				try {
				agentEmail = runnerClassObject.getEmailfromText(emailText);
				}
				catch(Exception e) {}
				System.out.println("Agent Email ="+agentEmail);
				
				Cell cell_ManagerEmail = row_Email.getCell(4);
				String emailText_Manager = cell_ManagerEmail.toString();
				System.out.println(emailText);
				String managerEmail=null ;
				try {
				 managerEmail = runnerClassObject.getEmailfromText(emailText_Manager);
				}catch(Exception e) {}
					System.out.println("Agent Email ="+agentEmail);
				mailObject.sendFileToMail(agentName,agentEmail,managerEmail);
				
				Row row_name = sheet.getRow(i);
				Cell cell_name = row_name.createCell(2);
				cell_name.setCellValue("Completed");
				agentEmailFileOutputStream = new FileOutputStream(agentEmails);
				workbook.write(agentEmailFileOutputStream);
			  }
			}catch(Exception e1)
				{
					Row row_name = sheet.getRow(i);
					Cell cell_name = row_name.createCell(5);
					cell_name.setCellValue("NA");
					agentEmailFileOutputStream = new FileOutputStream(agentEmails);
					workbook.write(agentEmailFileOutputStream);
				}
			}
			catch(Exception e)
			{
				//agentEmailFileOutputStream = new FileOutputStream(agentEmails);
				
				workbook.close();
			}
			
			}
			}
			
			//workbook.write(agentEmailFileOutputStream);
		}
			workbook.close();
			/*
			Row row = sheet.getRow(i+1);
			Cell cell = row.getCell(1);
			String emailText = cell.toString();
			System.out.println(emailText);
			runnerClassObject.getEmailfromText(emailText);
			
			*/
		for(int i=0;i<emailsList.size();i++)
		{
			System.out.println(emailsList.get(i));
		}
	}
	public String getEmailfromText(String email)
	{
		/*Matcher m = Pattern.compile("[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\\.[a-zA-Z0-9-.]+").matcher(email);
		System.out.println(m.matches());
	    while (m.find()) 
	    {
	    	
	        System.out.println(m.group());
	        emailsList.add(m.group());
	    }
	    return m.group(0);*/
	    
		Pattern p = Pattern.compile("[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\\.[a-zA-Z0-9-.]+");
		Matcher m = p.matcher(email);
		//System.out.println(m.matches());
		//m.matches();
		if (m.find()) {
		    String urlStr = m.group(0);
		    System.out.println(urlStr);
		}
	    return m.group(0);
	    
	     // System.out.println(m);
		//return m.group(0);
	}
}
