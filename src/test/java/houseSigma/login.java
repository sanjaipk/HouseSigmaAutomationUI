/**
 * 
 */
package houseSigma;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;
import java.util.concurrent.TimeUnit;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.json.JSONArray;
import org.json.JSONException;
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
	static Object[][] houseD = new Object[200][50];
	static Map<String,String> Heads = new LinkedHashMap<String,String>();
	static Map<String,String> URIs = new LinkedHashMap<String,String>();
	static List<String> MLSNumber = new ArrayList<String>();
	/**
	 * @param args
	 * @throws InterruptedException
	 * @throws IOException
	 */
	public static void main(String[] args) throws InterruptedException, IOException {

		URIs.put("bestschool","https://housesigma.com/web/en/recommend/more/bestschool");
		URIs.put("rental","https://housesigma.com/web/en/recommend/more/bestrental");
		URIs.put("pricegrowth","https://housesigma.com/web/en/recommend/more/pricegrowth");
		URIs.put("sold","https://housesigma.com/web/en/recommend/more/justsold");

		Heads.put("list_status@status","String");
		Heads.put("id_listing", "String");
		Heads.put("ml_num_merge","String");
		Heads.put("seo_suffix","String");
		Heads.put("address","String");
		Heads.put("house_type_name","String");
		Heads.put("house_style","String");
		Heads.put("community_name","String");
		Heads.put("province","String");
		Heads.put("municipality_name","String");
		
		Heads.put("price","String");
		Heads.put("price_sold","String");
		Heads.put("date_added","String");
		Heads.put("date_end","String");
	
		Heads.put("price_int","Int");
		Heads.put("price_sold_int","Int");
		
		Heads.put("bedroom","Int");
		Heads.put("bedroom_plus","Int");
		Heads.put("washroom","Int");
		
		Heads.put("list_days","Int");
		
		Heads.put("analytics@estimate_price","String");
		Heads.put("text@rooms_long","String");
		Heads.put("house_area@estimate","Int");
		Heads.put("land@depth","Int");
		Heads.put("land@front","Int");
		Heads.put("map@lon","Int");
		Heads.put("map@lat","Int");
		
		
		
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
			//recomm = "https://housesigma.com/web/en/recommend/more/justsold";
			driver.get("https://housesigma.com/web/en/user/watched");
			String title = driver.getTitle();
			Thread.sleep(2000);
			driver.findElement(By.xpath("//div[@id='tab-phone']")).click();
			Thread.sleep(2000);
			driver.findElement(By.xpath("//div[@id='pane-phone']//input[@name='account']")).sendKeys("4168791456");
			driver.findElement(By.xpath("//div[@id='pane-phone']//input[@type='password']")).sendKeys("Bar4sanjai!");
			driver.findElement(By.xpath("//button/span[text()=\"Sign-In\"]/parent::button")).click();
			Thread.sleep(4000);
			
			for (Iterator<Entry<String, String>> iterator2 = URIs.entrySet().iterator(); iterator2.hasNext();) {
				Entry<String, String> setting2 = iterator2.next();
				String key2 = setting2.getKey();
				String val2 = setting2.getValue();
				
				driver.get(val2);
				Thread.sleep(10000);
				
				JavascriptExecutor js = (JavascriptExecutor) driver;
				for (int i = 0; i < 17; i++) {
					js.executeScript("window.scrollTo(0, document.body.scrollHeight)");
					Thread.sleep(3000);
				}
			}
			

			System.out.println("Started processing");
			
			
			Thread.sleep(6000);
			
			File myObj = new File("test.har");
			myObj.createNewFile();
			
			
			List<HarEntry> harColl = bmps.getHar().getLog().getEntries();
			int j = 0;
			
			writeHouseDetail(j, null, true);
			j++;
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
					
					for (int i =0; i < jarr.length(); i++) {
						jobj = jarr.getJSONObject(i);
						String MLSNumberCurr = jobj.getString("ml_num_merge");
						if(MLSNumber.contains(MLSNumberCurr)) {
							continue; //avoid duplicates in excel
						} else {
							MLSNumber.add(MLSNumberCurr);
						}
						System.out.println(jobj.toString());
						writeHouseDetail(j, jobj, false);
						j++;
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

	private static void writeHouseDetail(int j, JSONObject jobj, boolean addHeader) {
		
			int k=0;
			for (Iterator<Entry<String, String>> iterator = Heads.entrySet().iterator(); iterator.hasNext();) {
				Entry<String, String> setting = iterator.next();
				String key = setting.getKey();
				String val = setting.getValue();
				if(addHeader) { 
					if(key.contains("@")) {
						try {
							String[] arrKeys = key.split("@", 2);
							System.out.println("test - "+arrKeys);
								houseD[j][k]= arrKeys[1];
						} catch (JSONException e) {
							e.printStackTrace();
						}
					} else {
						houseD[j][k]= key;
					}
				} else {
					if(key.contains("@")) {
						try {
							String[] arrKeys = key.split("@", 2);
							houseD[j][k] = val.equalsIgnoreCase("String")? jobj.getJSONObject(arrKeys[0]).getString(arrKeys[1]) : jobj.getJSONObject(arrKeys[0]).getInt(arrKeys[1]);
						} catch (JSONException e) {
							e.printStackTrace();
						}
					} else {
						try {
							if(key.equalsIgnoreCase("id_listing")) {
								houseD[j][k] =  "https://housesigma.com/web/en/house/"+ jobj.getString(key);
							} else {
								houseD[j][k] = val.equalsIgnoreCase("String")? jobj.getString(key) : jobj.getInt(key);
							}
						} catch (JSONException e) {
							e.printStackTrace();
						}
					}
				}
				k++;
			}
		
//		houseD[j][0] = jobj.getInt("date_start_days");
//		
//		houseD[j][2] = jobj.getInt("date_start_days");
//		houseD[j][3] = jobj.getString("price");
//		houseD[j][4] = jobj.getInt("date_added_days");
//		houseD[j][5] = jobj.getInt("price_int");
//		houseD[j][6] = jobj.getInt("date_start_month");
//		houseD[j][7] = jobj.getInt("list_days");
//		houseD[j][8] = jobj.getString("date_update");
//		houseD[j][9] = jobj.getString("community_name");
//		houseD[j][10] = jobj.getString("province");
//		houseD[j][11] = jobj.getString("price_abbr");
//		houseD[j][11] = jobj.getString("dom_long");
//		try {
//		houseD[j][12] = jobj.getJSONObject("land").getInt("depth");
//		houseD[j][13] = jobj.getJSONObject("land").getInt("front");
//		} catch (Exception ex) {
//			System.out.println(ex.getMessage());
//		}
//		houseD[j][14] = jobj.getString("municipality_name");
//		houseD[j][15] = jobj.getString("house_style");
//		houseD[j][16] = jobj.getString("house_type_name");
	}
	
	
}