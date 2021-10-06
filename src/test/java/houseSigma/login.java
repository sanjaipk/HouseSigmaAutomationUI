/**
 * 
 */
package houseSigma;

import io.github.bonigarcia.wdm.WebDriverManager;
import net.lightbody.bmp.BrowserMobProxyServer;
import net.lightbody.bmp.client.ClientUtil;
import net.lightbody.bmp.core.har.HarEntry;
import net.lightbody.bmp.proxy.CaptureType;

import java.io.File;
import java.io.IOException;
import java.util.List;
import java.util.concurrent.TimeUnit;

import org.openqa.selenium.By;
import org.openqa.selenium.Capabilities;
import org.openqa.selenium.Proxy;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.remote.CapabilityType;
import org.openqa.selenium.remote.DesiredCapabilities;

/**
 * @author m_166894
 *
 */
public class login {

	/**
	 * @param args
	 * @throws InterruptedException
	 * @throws IOException
	 */
	public static void main(String[] args) throws InterruptedException, IOException {
		ChromeDriver driver = null;
		BrowserMobProxyServer bmps = null;
		try {
			WebDriverManager.chromedriver().setup();

			bmps = new BrowserMobProxyServer();
			bmps.setTrustAllServers(true);
			bmps.start(0);
			bmps.setHarCaptureTypes(CaptureType.getAllContentCaptureTypes());
			bmps.enableHarCaptureTypes(CaptureType.REQUEST_CONTENT, CaptureType.RESPONSE_CONTENT);
	       
			Proxy seliniumProxy = ClientUtil.createSeleniumProxy(bmps);
			//seliniumProxy.setHttpProxy("localhost:"+bmps.getPort());
			//seliniumProxy.setSslProxy("localhost:"+bmps.getPort());
			ChromeOptions options = new ChromeOptions();
			options.addArguments("--ignore-certificate-errors");
			options.setCapability(CapabilityType.ACCEPT_SSL_CERTS, true);
			options.setCapability(CapabilityType.PROXY, seliniumProxy);
			driver = new ChromeDriver(options);
			driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
			driver.manage().window().maximize();
			bmps.newHar("housesigmaoutput");
			
			driver.get("https://housesigma.com/web/en/user/watched");
			String title = driver.getTitle();
			Thread.sleep(2000);
			driver.findElement(By.xpath("//div[@id='tab-phone']")).click();
			Thread.sleep(2000);
			driver.findElement(By.xpath("//div[@id='pane-phone']//input[@name='account']")).sendKeys("4168791456");
			driver.findElement(By.xpath("//div[@id='pane-phone']//input[@type='password']")).sendKeys("Bar4sanjai!");
			driver.findElement(By.xpath("//button/span[text()=\"Sign-In\"]/parent::button")).click();
			Thread.sleep(3000);

			File myObj = new File("test.har");
			myObj.createNewFile();
			
			List<HarEntry> harColl = bmps.getHar().getLog().getEntries();
			for (HarEntry harnetry : harColl) {
				String currentURL = harnetry.getRequest().getUrl();
				String currentAPI = currentURL.substring(currentURL.lastIndexOf("/") + 1, currentURL.length()).trim();
				boolean isGetHTTPCall = harnetry.getRequest().getMethod().toString().toUpperCase().contains("GET");
				boolean isPostHTTPCall = harnetry.getRequest().getMethod().toString().toUpperCase().contains("POST");
				boolean isMemberLogin = harnetry.getRequest().getUrl().toLowerCase().contains("memberlogin");
				
			}
			//driver.quit();
			bmps.stop();
		} catch (Exception e) {
			e.printStackTrace();
		} finally {
			if (bmps != null) {
				if(!bmps.isStopped()){
					bmps.stop();
				}
			}
			if (driver != null) {
				//driver.quit();
			}
		}

	}

}
