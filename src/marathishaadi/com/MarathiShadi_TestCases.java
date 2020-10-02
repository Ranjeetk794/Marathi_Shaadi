package marathishaadi.com;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;

public class MarathiShadi_TestCases {
	public static void main(String args[]) throws Exception {
		// Expected Community name variable
		String expCommunity_name = "Marathi";
		ExcelData excell = new ExcelData("./data/marathishadi.xlsx");
		WebDriver_Lib wlb = new WebDriver_Lib();
		// System.out.println(excell.getData("Data",1, 3));
		System.setProperty("webdriver.chrome.driver", "./browser/chromedriver.exe");
		WebDriver driver = new ChromeDriver();
		// Open the community domain
		driver.get(excell.getPropetyKeyvalue("Url"));
		// To maximize the browser
		System.out.println(driver.getTitle());

		driver.manage().window().maximize();
		// implicit wait
		wlb.waitforPageLoad(driver);
		// Click on let's begin button
		driver.findElement(By.xpath("//button[@data-testid='lets_begin']")).click();
		// Enter the data in Email Text
		driver.findElement(By.name("email")).sendKeys(excell.getData("Data", 1, 0));
		// Enter the data in Password Text
		driver.findElement(By.name("password1")).sendKeys(excell.getData("Data", 0, 1));
		// Hard coded wait
		Thread.sleep(1000);
		// Click on doropdown root
		driver.findElement(By.xpath("//div[contains(@class,'Dropdown-root')]")).click();
		// Select daughter
		driver.findElement(By.xpath("//div[text()='Daughter']")).click();
		// Click on next button
		driver.findElement(By.xpath("//button[@data-testid='next_button']")).click();
		// Enter first Name Text
		driver.findElement(By.xpath("//input[@data-testid='first_name']")).sendKeys(excell.getData("Data", 1, 2));
		// Enter Last Name Text
		driver.findElement(By.xpath("//input[@data-testid='last_name']")).sendKeys(excell.getData("Data", 1, 3));
		// Select Drop Down
		driver.findElement(By.xpath("//div[@class='Dropdown-control day_selector']")).click();
		// Select Date
		driver.findElement(By.xpath("//div[@class='Dropdown-option'][4]")).click();
		// Select Drop down Month
		driver.findElement(By.xpath("//div[@class='Dropdown-control month_selector']")).click();
		// Click month
		driver.findElement(By.xpath("//div[text()='Feb']")).click();
		// Select year drop down year
		driver.findElement(By.xpath("//div[@class='Dropdown-control year_selector']")).click();
		// Thread.sleep(2000);
		driver.findElement(By.xpath("//div[text()='2000']")).click();
		// Click Community
		driver.findElement(By.xpath("//div[@class='Dropdown-control community_selector']")).click();
		driver.findElement(By.xpath("//div[text()='Hindu']")).click();
		// Verify community tounge
		WebElement tounge = driver.findElement(By.xpath("//div[text()='Marathi']"));

		System.out.println("Mother Tongue:- " + tounge.getText());
		// Match with community name and mother tongue is the same
		if (expCommunity_name.equals(tounge.getText())) {
			System.out.println("Community name and mother tongue is the same");
		} else {
			System.out.println("Community name and mother tongue is the not same");
		}
		;
		driver.findElement(By.xpath("//div[@class='Dropdown-control countryofresidence_selector']")).click();
		// Select country
		driver.findElement(By.xpath("//div[text()='India']")).click();
		// Close the browser
		driver.close();

	}

}
