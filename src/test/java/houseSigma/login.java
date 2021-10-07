/**
 * 
 */
package houseSigma;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;
import java.util.concurrent.TimeUnit;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.json.JSONArray;
import org.json.JSONObject;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.Proxy;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.remote.CapabilityType;

import io.github.bonigarcia.wdm.WebDriverManager;
import net.lightbody.bmp.BrowserMobProxyServer;
import net.lightbody.bmp.client.ClientUtil;
import net.lightbody.bmp.core.har.HarEntry;
import net.lightbody.bmp.proxy.CaptureType;

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
			seliniumProxy.setHttpProxy("localhost:" + bmps.getPort());
			seliniumProxy.setSslProxy("localhost:" + bmps.getPort());
			ChromeOptions options = new ChromeOptions();
			options.addArguments("--ignore-certificate-errors");
			options.setCapability(CapabilityType.ACCEPT_SSL_CERTS, true);
			options.setCapability(CapabilityType.PROXY, seliniumProxy);
			driver = new ChromeDriver(options);
			driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
			driver.manage().window().maximize();
			bmps.newHar("housesigmaoutput");

			String recomm = "https://housesigma.com/web/en/recommend/more/bestrental";
			driver.get("https://housesigma.com/web/en/user/watched");
			String title = driver.getTitle();
			Thread.sleep(2000);
			driver.findElement(By.xpath("//div[@id='tab-phone']")).click();
			Thread.sleep(2000);
			driver.findElement(By.xpath("//div[@id='pane-phone']//input[@name='account']")).sendKeys("4168791456");
			driver.findElement(By.xpath("//div[@id='pane-phone']//input[@type='password']")).sendKeys("Bar4sanjai!");
			driver.findElement(By.xpath("//button/span[text()=\"Sign-In\"]/parent::button")).click();
			Thread.sleep(4000);
			driver.get(recomm);
			Thread.sleep(10000);
			
			JavascriptExecutor js = (JavascriptExecutor) driver;
			for (int i = 0; i < 17; i++) {
				js.executeScript("window.scrollTo(0, document.body.scrollHeight)");
				Thread.sleep(3000);
			}

			System.out.println("Started processing");
			Thread.sleep(6000);
			
			File myObj = new File("test.har");
			myObj.createNewFile();
			
			Object[][] houseD = new Object[200][17];
			List<HarEntry> harColl = bmps.getHar().getLog().getEntries();
			for (HarEntry harnetry : harColl) {
				String currentURL = harnetry.getRequest().getUrl();
				String currentAPI = currentURL.substring(currentURL.lastIndexOf("/") + 1, currentURL.length()).trim();
				boolean isGetHTTPCall = harnetry.getRequest().getMethod().toString().toUpperCase().contains("GET");
				boolean isPostHTTPCall = harnetry.getRequest().getMethod().toString().toUpperCase().contains("POST");
				boolean isMemberLogin = harnetry.getRequest().getUrl().toLowerCase().contains("homepage/recommendlist");
				if (isMemberLogin) {
					String currentResponse = harnetry.getResponse().getContent().getText();
					JSONObject jobj = new JSONObject(currentResponse);
					jobj = jobj.getJSONObject("data");
					JSONArray jarr = jobj.getJSONArray("list");
					
					for (int i = 1; i < jarr.length(); i++) {
						jobj = jarr.getJSONObject(i);
						System.out.println(jobj.toString());
						houseD[i - 1][0] = jobj.getInt("date_start_days");
						houseD[i - 1][1] = jobj.getString("ml_num_merge");
						houseD[i - 1][2] = jobj.getInt("date_start_days");
						houseD[i - 1][3] = jobj.getString("price");
						houseD[i - 1][4] = jobj.getInt("date_added_days");
						houseD[i - 1][5] = jobj.getInt("price_int");
						houseD[i - 1][6] = jobj.getInt("date_start_month");
						houseD[i - 1][7] = jobj.getInt("list_days");
						houseD[i - 1][8] = jobj.getString("date_update");
						houseD[i - 1][9] = jobj.getString("community_name");
						houseD[i - 1][10] = jobj.getString("province");
						houseD[i - 1][11] = jobj.getString("price_abbr");
						houseD[i - 1][12] = jobj.getJSONObject("land").getInt("depth");
						houseD[i - 1][13] = jobj.getJSONObject("land").getInt("front");
						houseD[i - 1][14] = jobj.getString("municipality_name");
						houseD[i - 1][15] = jobj.getString("house_style");
						houseD[i - 1][16] = jobj.getString("house_type_name");
					}
				}
			}
			System.out.println("excelling");
			String excelFilePath = "/Users/m_166894/Desktop/JavaBooks.xlsx";
			try {
				FileInputStream inputStream = new FileInputStream(new File(excelFilePath));
				Workbook workbook = WorkbookFactory.create(inputStream);
				Sheet sheet = workbook.getSheetAt(0);
				int rowCount = sheet.getLastRowNum();
				for (Object[] aBook : houseD) {
					Row row = sheet.createRow(++rowCount);
					int columnCount = 0;
					Cell cell = row.createCell(columnCount);
					cell.setCellValue(rowCount);
					for (Object field : aBook) {
						cell = row.createCell(++columnCount);
						if (field instanceof String) {
							cell.setCellValue((String) field);
						} else if (field instanceof Integer) {
							cell.setCellValue((Integer) field);
						}
					}
				}
				inputStream.close();
				FileOutputStream outputStream = new FileOutputStream(excelFilePath);
				workbook.write(outputStream);
				workbook.close();
				outputStream.close();
			} catch (Exception ex) {
				ex.printStackTrace();
			}
			driver.quit();
			bmps.stop();
		} catch (Exception e) {
			e.printStackTrace();
		} finally {
			if (bmps != null) {
				if (!bmps.isStopped()) {
					bmps.stop();
				}
			}
			if (driver != null) {
				driver.quit();
			}
		}
	}
}