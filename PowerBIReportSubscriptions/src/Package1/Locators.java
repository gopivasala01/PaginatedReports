package Package1;

import org.openqa.selenium.By;

public class Locators 
{
	public static By userName = By.name("loginfmt");
	public static By userName2 = By.id("email");
	public static By loginButton = By.id("idSIButton9");
	public static By loginButton2 = By.id("submitBtn");
	public static By password = By.name("passwd");
	public static By signInButton = By.id("idSIButton9");
	public static By workspacesButton = By.xpath("//*[@class='workspacesPaneExpander switcher']");
	public static By workspacesList = By.xpath("//*[@class='workspaceName']");
	public static By reportList = By.xpath("//*[@id='artifactContentView']/div[1]/div/div[2]/span/a");
	public static By leasingAgentDropdown = By.name("VWUnitsLeasesMarketedManagedPrimaryLeasingAgent-visible");
	public static By leasingAgentsList = By.xpath("//*[@role='menuitemcheckbox']");
	public static By leasingAgentsListCheckboxes = By.xpath("//*[@role='menuitemcheckbox']/input");
	public static By clickOutsideOfTheReport = By.id("AnaheimHost");
	public static By viewReportButton = By.name("Parameter_ViewReport");
	public static By reportFrame = By.xpath("//*[@id='content']/descendant::iframe");
	public static By subscribeButton = By.xpath("//*[@title='Subscribe']");
	public static By addNewSubscription = By.xpath("//*[@class='pbi-side-pane paneLayout']/descendant::button[2]");
	public static By formatType = By.xpath("//*[@id='undefined']/section[2]/fieldset[1]/select");
	public static By excelFormat = By.xpath("//*[@aria-labelledby='export-type-label-0']/option[1]");
	public static By enterEmail = By.xpath("//*[@placeholder='Enter email addresses']");
	public static By subject = By.id("subject-input-0");
	public static By permissionsToAccessPowerBI = By.xpath("//*[@name='subscriptionForm']/subscription-pane-subscription-content/section[2]/div[2]/descendant::input");
	//public static By permissionsToAccessPowerBI = By.id("mat-checkbox-2");
	public static By linkToReportInPowerBI = By.xpath("//*[@name='subscriptionForm']/subscription-pane-subscription-content/section[2]/div[3]/descendant::input");
	public static By saveAndClose = By.xpath("//*[@title='Save and close']");
	public static By exportButton = By.xpath("//*[@title='Export']");
	public static By excelButton = By.xpath("(//*[@title='Microsoft Excel (.xlsx)'])[2]");
	public static By reportExists = By.xpath("//*[@class='Page']");
	public static By spinner = By.xpath("//*[@class='app-spinner-overlay']");
	
	
	public static By outlookEmail = By.xpath("//*[@aria-labelledby='label-username']");
	public static By outlookPassword = By.xpath("//*[@aria-labelledby='label-password']");
	public static By outlookLoginButton = By.id("submitBtn");
	
	
	
}
