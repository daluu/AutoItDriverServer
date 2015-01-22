import org.junit.*;
import org.openqa.selenium.*;
import org.openqa.selenium.remote.DesiredCapabilities;
import org.openqa.selenium.remote.RemoteWebDriver;
import java.net.URL;

public class CalculatorTest {
	
	public WebDriver driver;
	public DesiredCapabilities capabilities;

	@Before
	public void setUp() throws Exception {
		capabilities = new DesiredCapabilities();
		capabilities.setCapability("browserName", "AutoIt");
		driver = new RemoteWebDriver(new URL("http://localhost:4723/wd/hub" ), capabilities);
	}

	@After
	public void tearDown() throws Exception {
		driver.quit();
	}

	@Test
	public void test() throws Exception{
		// demo adapted from 
		// http://www.joecolantonio.com/2014/07/02/selenium-autoit-how-to-automate-non-browser-based-functionality/
		driver.get("calc.exe");
		driver.switchTo().window("Calculator");
		Thread.sleep(1000);
		driver.findElement(By.id("133")).click(); // 3
		Thread.sleep(1000);
		driver.findElement(By.id("93")).click(); // +
		Thread.sleep(1000);
		driver.findElement(By.id("133")).click(); // 3
		Thread.sleep(1000);
		driver.findElement(By.id("121")).click(); // =
		Thread.sleep(1000);
		Assert.assertEquals("3 + 3 did not produce 6 as expected.", "6", driver.findElement(By.id("150")).getText());
		driver.findElement(By.id("81")).click(); // Clear "C" button
		Thread.sleep(1000);
		driver.close();
	}

}
