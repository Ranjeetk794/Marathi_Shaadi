package marathishaadi.com;

import java.util.concurrent.TimeUnit;

import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;

public class WebDriver_Lib {
	/**
	 * its used to implicitly wait target for entire page load before perform action
	 * on element maximum time unit 15 second
	 * 
	 * @param Driver
	 */
	public void waitforPageLoad(WebDriver Driver) {
		Driver.manage().timeouts().implicitlyWait(15, TimeUnit.SECONDS);
	}

	

	/**
	 * its used to explicitly wait for Target specific element maximum time unit 15
	 * seconds
	 * 
	 * @param Driver
	 * @param element
	 */
	public void waitforElementPresent(WebDriver driver,WebElement element) {
		WebDriverWait wait = new WebDriverWait(driver, 15);
		wait.until(ExpectedConditions.elementToBeClickable(element));
	}

}
