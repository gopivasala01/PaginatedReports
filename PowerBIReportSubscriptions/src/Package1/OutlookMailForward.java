package Package1;

public class OutlookMailForward 
{
	OutlookMailForward()
	{
		RunnerClass.driver.get(AppConfig.outlookURL);
		RunnerClass.driver.manage().window().maximize();
	}
	public static void loginToMail()
	{
		RunnerClass.driver.findElement(Locators.outlookEmail).sendKeys("gopi.v@beetlerim.com");
		RunnerClass.driver.findElement(Locators.outlookPassword).sendKeys("Welcome@123");
		RunnerClass.driver.findElement(Locators.outlookLoginButton).click();
	}

}
