/**
 * 
 */
package houseSigma;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.PrintWriter;
import java.math.BigInteger;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;
import java.util.Scanner;
import java.util.concurrent.TimeUnit;
import java.util.stream.Collectors;

import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.json.JSONArray;
import org.json.JSONException;
import org.json.JSONObject;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.Proxy;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.remote.CapabilityType;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPageMar;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSectPr;

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
	static Map<String, String> Heads = new LinkedHashMap<String, String>();
	static Map<String, String> URIs = new LinkedHashMap<String, String>();
	static List<String> MLSNumber = new ArrayList<String>();

	/**
	 * @param args
	 * @throws InterruptedException
	 * @throws IOException
	 */
	public static void main(String[] args) throws InterruptedException, IOException {

//		URIs.put("bestschool","https://housesigma.com/web/en/recommend/more/bestschool");
//		URIs.put("rental","https://housesigma.com/web/en/recommend/more/bestrental");
//		URIs.put("pricegrowth","https://housesigma.com/web/en/recommend/more/pricegrowth");
//		URIs.put("sold","https://housesigma.com/web/en/recommend/more/justsold");
		URIs.put("watched", "https://housesigma.com/web/en/user/watched");
		URIs.put("witbhyMap",
				"https://housesigma.com/web/en/map?zoom=14&center=%7B%22lat%22%3A43.899803988558915,%22lng%22%3A-78.9393230318092%7D");
		URIs.put("scbroughMap",
				"https://housesigma.com/web/en/map?zoom=13&center=%7B%22lat%22%3A43.73291363273812,%22lng%22%3A-79.25255175679922%7D");
		URIs.put("testMap",
				"https://housesigma.com/web/en/map?zoom=17&center=%7B%22lat%22%3A43.843263545576264,%22lng%22%3A-79.00349828076288%7D");
		
		Heads.put("list_status@status", "String");
		Heads.put("id_listing", "String");
		Heads.put("ml_num_merge", "String");
		Heads.put("seo_suffix", "String");
		Heads.put("address", "String");
		Heads.put("house_type_name", "String");
		Heads.put("house_style", "String");
		Heads.put("community_name", "String");
		Heads.put("province", "String");
		Heads.put("municipality_name", "String");

		Heads.put("price", "String");
		Heads.put("price_sold", "String");
		Heads.put("date_added", "String");
		Heads.put("date_end", "String");

		Heads.put("price_int", "Int");
		Heads.put("price_sold_int", "Int");

		Heads.put("bedroom", "Int");
		Heads.put("bedroom_plus", "Int");
		Heads.put("washroom", "Int");

		Heads.put("list_days", "Int");

		Heads.put("analytics@estimate_price", "String");
		Heads.put("text@rooms_long", "String");
		Heads.put("house_area@estimate", "Int");
		Heads.put("land@depth", "Int");
		Heads.put("land@front", "Int");
		Heads.put("map@lon", "Int");
		Heads.put("map@lat", "Int");

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
			String mapbased = "";
			// recomm = "https://housesigma.com/web/en/recommend/more/justsold";
			String watched = "https://housesigma.com/web/en/user/watched";
			driver.get(watched);
			String title = driver.getTitle();
			Thread.sleep(2000);
			driver.findElement(By.xpath("//div[@id='tab-phone']")).click();
			Thread.sleep(2000);
			driver.findElement(By.xpath("//div[@id='pane-phone']//input[@name='account']")).sendKeys("4168791456");
			driver.findElement(By.xpath("//div[@id='pane-phone']//input[@type='password']")).sendKeys("Bar4sanjai!");
			driver.findElement(By.xpath("//button/span[text()=\"Sign-In\"]/parent::button")).click();
			Thread.sleep(4000);
			
			processMLSFRomFile(driver);
//			saveList(MLSNumber);
			processDetails(bmps.getHar().getLog().getEntries());
			driver.get(URIs.get("scbroughMap"));
//			for (Iterator<Entry<String, String>> iterator2 = URIs.entrySet().iterator(); iterator2.hasNext();) {
//				Entry<String, String> setting2 = iterator2.next();
//				String key2 = setting2.getKey();
//				String val2 = setting2.getValue();
//
//				driver.get(val2);
//				Thread.sleep(10000);
//
//				JavascriptExecutor js = (JavascriptExecutor) driver;
//				for (int i = 0; i < 17; i++) {
//					js.executeScript("window.scrollTo(0, document.body.scrollHeight)");
//					Thread.sleep(3000);
//				}
//			}

			System.out.println("Started processing");

			Thread.sleep(6000);

			File myObj = new File("test.har");
			myObj.createNewFile();

			List<HarEntry> harColl = bmps.getHar().getLog().getEntries();
			int j = 0;

			writeHouseDetail(j, null, true);
			j++;
			boolean isMapSearchProcess = false;
			for (HarEntry harnetry : harColl) {
				String currentURL = harnetry.getRequest().getUrl();
				String currentAPI = currentURL.substring(currentURL.lastIndexOf("/") + 1, currentURL.length()).trim();
				boolean isGetHTTPCall = harnetry.getRequest().getMethod().toString().toUpperCase().contains("GET");
				boolean isPostHTTPCall = harnetry.getRequest().getMethod().toString().toUpperCase().contains("POST");
				boolean isMemberLogin = harnetry.getRequest().getUrl().toLowerCase().contains("/mapsearchv2/listing2");
				boolean isMapSearch = harnetry.getRequest().getUrl().toLowerCase().contains("/mapsearchv2/listing2");
				// "homepage/recommendlist");
				// "/watch/list"
				if (isMemberLogin) {
					String currentResponse = harnetry.getResponse().getContent().getText();
					JSONObject jobj = new JSONObject(currentResponse);
					jobj = jobj.getJSONObject("data");
					JSONArray jarr = jobj.getJSONArray("list");// houselist

					if (isMapSearch) {
						isMapSearchProcess = true;
						// get ids
						for (int i = 0; i < jarr.length(); i++) {
							jobj = jarr.getJSONObject(i);
							JSONArray mlsids = jobj.getJSONArray("ids");
							for (int k = 0; k < mlsids.length(); k++) {
								String mlsid = mlsids.getString(k);
								System.out.println("Loading " + k + " out of " + mlsids.length() + " - " + mlsid);
								// open new browser and then process the detail
								// printMLSDetail(mlsid);
								if (MLSNumber.contains(mlsid)) {
									continue; // avoid duplicates in excel
								} else {
									MLSNumber.add(mlsid);
//									driver.get("https://housesigma.com/web/en/house/" + mlsid);
//									Thread.sleep(10000);
//
//									JavascriptExecutor js = (JavascriptExecutor) driver;
//									for (int l = 0; l < 3; l++) {
//										js.executeScript("window.scrollTo(0, document.body.scrollHeight)");
//										Thread.sleep(1000);
//									}
								}

							}
						}
						saveList(MLSNumber);
						
						for (String mlsnum : MLSNumber) {
							driver.get("https://housesigma.com/web/en/house/" + mlsnum);
							Thread.sleep(1000);
						}
						processDetails(bmps.getHar().getLog().getEntries());
					}

//					for (int i =0; i < jarr.length(); i++) {
//						jobj = jarr.getJSONObject(i);
//						String MLSNumberCurr = jobj.getString("ml_num_merge");
//						if(MLSNumber.contains(MLSNumberCurr)) {
//							continue; //avoid duplicates in excel
//						} else {
//							MLSNumber.add(MLSNumberCurr);
//						}
//						System.out.println(jobj.toString());
//						writeHouseDetail(j, jobj, false);
//						j++;
//					}
				}
			}

			System.out.println("excelling");
//			String excelFilePath = "/Users/m_166894/Desktop/JavaBooks.xlsx";
//			try {
//				FileInputStream inputStream = new FileInputStream(new File(excelFilePath));
//				Workbook workbook = WorkbookFactory.create(inputStream);
//				Sheet sheet = workbook.getSheetAt(0);
//				int rowCount = sheet.getLastRowNum();
//				for (Object[] aBook : houseD) {
//					Row row = sheet.createRow(++rowCount);
//					int columnCount = 0;
//					Cell cell = row.createCell(columnCount);
//					cell.setCellValue(rowCount);
//					for (Object field : aBook) {
//						cell = row.createCell(++columnCount);
//						if (field instanceof String) {
//							cell.setCellValue((String) field);
//						} else if (field instanceof Integer) {
//							cell.setCellValue((Integer) field);
//						}
//					}
//				}
//				inputStream.close();
//				FileOutputStream outputStream = new FileOutputStream(excelFilePath);
//				workbook.write(outputStream);
//				workbook.close();
//				outputStream.close();
//			} catch (Exception ex) {
//				ex.printStackTrace();
//			}
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

	private static void saveList(List<String> mLSNumber2) throws FileNotFoundException, IOException {
		File file = new File("output.txt");

		try (PrintWriter pw = new PrintWriter(new FileOutputStream(file))) {
			int datList = mLSNumber2.size();

			for (String s : mLSNumber2) {
				pw.println(s);
			}
		} catch (Exception e) {
			e.printStackTrace();
		}

	}

	private static void readList() throws FileNotFoundException {
		Scanner s = new Scanner(new File("output.txt"));
		MLSNumber = new ArrayList<String>();
		while (s.hasNext()) {
			MLSNumber.add(s.next());
		}
		s.close();
	}
	
	private static void processMLSFRomFile(ChromeDriver driver) throws InterruptedException, FileNotFoundException {
		readList();
		for (String mlsid : MLSNumber) {
			driver.get("https://housesigma.com/web/en/house/" + mlsid);
			Thread.sleep(500);
//
//			JavascriptExecutor js = (JavascriptExecutor) driver;
//			for (int l = 0; l < 3; l++) {
//				js.executeScript("window.scrollTo(0, document.body.scrollHeight)");
//				Thread.sleep(1000);
//			}
		}		
	}

	private static void processDetails(List<HarEntry> entries) throws FileNotFoundException {
		for (HarEntry harnetry : entries) {
			String currentURL = harnetry.getRequest().getUrl();
			String currentAPI = currentURL.substring(currentURL.lastIndexOf("/") + 1, currentURL.length()).trim();
			boolean isGetHTTPCall = harnetry.getRequest().getMethod().toString().toUpperCase().contains("GET");
			boolean isPostHTTPCall = harnetry.getRequest().getMethod().toString().toUpperCase().contains("POST");
			boolean isDetail = harnetry.getRequest().getUrl().toLowerCase().contains("/api/listing/info/detail");
			int FontSize = 8;
			if (isDetail && isPostHTTPCall) {
				try {
					String currentResponse = harnetry.getResponse().getContent().getText();
					JSONObject jobj = new JSONObject(currentResponse);
					jobj = jobj.getJSONObject("data");
					JSONObject house = jobj.getJSONObject("house");

					// search Params:
					int hprice = jobj.getJSONObject("house").getInt("price_int");
					String htype = jobj.getJSONObject("house").getString("house_type_name");
					float hlanddepth = jobj.getJSONObject("house").getJSONObject("land").getFloat("depth");
					float hlandfront = jobj.getJSONObject("house").getJSONObject("land").getFloat("front");
					boolean isSelected = false;

					if ((hprice < 1000000) && (htype.equalsIgnoreCase("Detached")) && (hlanddepth > 70.0)
							&& (hlandfront > 25.0)) {
						isSelected = true;
					}

					if (isSelected) {
						JSONObject ksy_facts = jobj.getJSONObject("key_facts");
						JSONObject prop_details = jobj.getJSONObject("property_detail");
						JSONArray rooms = null;
						JSONArray listing_history = null;
						JSONObject property_detail = null;
						JSONObject analytics = null;

						try {
							rooms = jobj.getJSONArray("rooms");
						} catch (Exception e) {
							System.out.println(e.getMessage());
						}
						try {
							listing_history = jobj.getJSONArray("listing_history");
						} catch (Exception e) {
							System.out.println(e.getMessage());
						}
						try {
							property_detail = jobj.getJSONObject("property_detail");
						} catch (Exception e) {
							System.out.println(e.getMessage());
						}
						try {
							analytics = jobj.getJSONObject("analytics");
						} catch (Exception e) {
							System.out.println(e.getMessage());
						}

						String fileName = house.getString("seo_suffix");

						System.out.println("Processing from har for " + fileName);
						XWPFDocument document = new XWPFDocument();
						try (FileOutputStream out = new FileOutputStream(new File(fileName + ".docx"))) {

							CTSectPr sectPr = document.getDocument().getBody().getSectPr();
							if (sectPr == null)
								sectPr = document.getDocument().getBody().addNewSectPr();
							CTPageMar pageMar = sectPr.getPgMar();
							if (pageMar == null)
								pageMar = sectPr.addNewPgMar();
							pageMar.setLeft(BigInteger.valueOf(720)); // 720 TWentieths of an Inch Point (Twips) =
																		// 720/20 = 36
																		// pt = 36/72 = 0.5"
							pageMar.setRight(BigInteger.valueOf(720));
							pageMar.setTop(BigInteger.valueOf(720));
							pageMar.setBottom(BigInteger.valueOf(720));
							pageMar.setFooter(BigInteger.valueOf(720));
							pageMar.setHeader(BigInteger.valueOf(720));
							pageMar.setGutter(BigInteger.valueOf(0));

							// write to first row, first column
							List<String> consol = new ArrayList<String>();
							for (Map.Entry<String, Object> ksy_fact : ksy_facts.toMap().entrySet()) {
								String key = ksy_fact.getKey();
								String value = ksy_fact.getValue() != null ? ksy_fact.getValue().toString() : "_";
								String keyval = key + "~" + value;
								if (consol.contains(keyval)) {
									continue; // avoid duplicates in excel
								} else {
									consol.add(keyval);
								}
							}
							for (Map.Entry<String, Object> house_fact : house.toMap().entrySet()) {
								String key = house_fact.getKey();
								String value = house_fact.getValue() != null ? house_fact.getValue().toString() : "_";
								String keyval = key + "~" + value;
								if (consol.contains(keyval)) {
									continue; // avoid duplicates in excel
								} else {
									consol.add(keyval);
								}
							}
							for (Map.Entry<String, Object> prop : prop_details.toMap().entrySet()) {
								Map<String, Object> tmp = (Map<String, Object>) prop.getValue();
								for (Map.Entry<String, Object> tmpprop : tmp.entrySet()) {
									if(tmpprop.getKey().equalsIgnoreCase("value")) {
										Map<String, Object> tmpp = (Map<String, Object>) prop.getValue();
										
										ArrayList<HashMap> arrValue = (ArrayList<HashMap>) tmpp.get("value");
										for (HashMap prop_val : arrValue) {
											String key = prop_val.get("name").toString();
											String value = prop_val.get("value").toString();
											String keyval2 = key + "~" + value;
											System.out.println(keyval2);
											if (consol.contains(keyval2)) {
												continue; // avoid duplicates in excel
											} else {
												consol.add(keyval2);
											}
										}
										
										
									}
								}
							}
							
							
							
							
//									XWPFParagraph p1 = row.getCell(0).getParagraphs().get(0);
//									p1.setAlignment(ParagraphAlignment.LEFT);
//									XWPFRun r1 = p1.createRun();
//									XWPFRun r12 = p1.createRun();
//									r1.setFontSize(8);
//									r12.setFontSize(8);
//									r1.setBold(true);
//									r1.setText(key);
//									// r1.addTab();
//									r1.setText(" :- ");
//									r12.setBold(false);
//									r12.setText(value);
//									r12.addBreak();
//							}
							
							List<String> consol_bigge = new ArrayList<String>();
							List<String> consol_small = new ArrayList<String>();
							consol_bigge = consol.stream().filter(s -> s.length()>=50).collect(Collectors.toList());
							consol_small = consol.stream().filter(s -> s.length()<50).collect(Collectors.toList());
							int splitNo = (int) Math.ceil(consol_small.size()/3.0) ;
							
							// Creating Table
							XWPFTable tab = document.createTable();
							tab.setWidth("100%");
							XWPFTableRow row = tab.getRow(0); // First row // Columns
							
							XWPFParagraph p1 = null;
							XWPFTableCell cell = row.getCell(0);
							
							for (int i = 0; i < consol_small.size(); i++) {
								String printable = consol_small.get(i);
								String[] parts = printable.split("~");
								if(i%splitNo == 0 && i !=0) {
									cell = row.addNewTableCell();
									cell.setText("");
								}
								p1 = cell.getParagraphs().get(0);
								p1.setAlignment(ParagraphAlignment.LEFT);
								XWPFRun r1 = p1.createRun();
								XWPFRun r12 = p1.createRun();
								r1.setFontSize(8);
								r12.setFontSize(8);
								String key = parts[0];
								String value = parts[1];
								r1.setBold(true);
								r1.setText(key);
								// r1.addTab();
								r1.setText(" :- ");
								r12.setBold(false);
								r12.setText(value);
								r12.addBreak();
								
							}
							
//							for (Map.Entry<String, Object> house_fact : house.toMap().entrySet()) {
//								XWPFParagraph p1 = row.getCell(1).getParagraphs().get(0);
//								p1.setAlignment(ParagraphAlignment.LEFT);
//								XWPFRun r1 = p1.createRun();
//								XWPFRun r12 = p1.createRun();
//								r1.setFontSize(8);
//								r12.setFontSize(8);
//								String key = house_fact.getKey();
//								String value = house_fact.getValue() != null ? house_fact.getValue().toString() : "";
//								r1.setBold(true);
//								r1.setText(key);
//								// r1.addTab();
//								r1.setText(" :- ");
//								r12.setBold(false);
//								r12.setText(value);
//								r12.addBreak();
//							}

//							XWPFTable table1 = document.createTable(1,1); // This is your row 1
//							XWPFTable table2 = document.createTable(1,3); // This is your row 2
//
//							// Now it's time to span each column of table1 and table2 to a span of your choice
//							// lets say 6 is the total span required assuming there's some row with 6 columns.
//
//							spanCellsAcrossRow(table1, 0, 0, 6,"sdfsdf");
//							spanCellsAcrossRow(table2, 0, 0, 2,"sdfsdf");
//							spanCellsAcrossRow(table2, 0, 1, 2,"sdfsdf");
//							spanCellsAcrossRow(table2, 0, 2, 2,"sdfsdf");

							row = tab.createRow(); // Second Row
							row.getCell(0).setText("");
							if (row.getCell(0).getCTTc().getTcPr() == null)
								row.getCell(0).getCTTc().addNewTcPr();
							if (row.getCell(0).getCTTc().getTcPr().getGridSpan() == null)
								row.getCell(0).getCTTc().getTcPr().addNewGridSpan();
							row.getCell(0).getCTTc().getTcPr().getGridSpan().setVal(BigInteger.valueOf((long) 3));

							for (int i = 0; i < consol_bigge.size(); i++) {
								String printable = consol_bigge.get(i);
								String[] parts = printable.split("~");
								p1 = row.getCell(0).getParagraphs().get(0);
								p1.setAlignment(ParagraphAlignment.LEFT);
								XWPFRun r1 = p1.createRun();
								XWPFRun r12 = p1.createRun();
								r1.setFontSize(8);
								r12.setFontSize(8);
								String key = parts[0];
								String value = parts[1];
								r1.setBold(true);
								r1.setText(key);
								// r1.addTab();
								r1.setText(" :- ");
								r12.setBold(false);
								r12.setText(value);
								r12.addBreak();
								
							}
							
							row = tab.createRow(); // Second Row
							row.getCell(0).setText("");
							if (row.getCell(0).getCTTc().getTcPr() == null)
								row.getCell(0).getCTTc().addNewTcPr();
							if (row.getCell(0).getCTTc().getTcPr().getGridSpan() == null)
								row.getCell(0).getCTTc().getTcPr().addNewGridSpan();
							row.getCell(0).getCTTc().getTcPr().getGridSpan().setVal(BigInteger.valueOf((long) 3));

							if (rooms != null) {
								for (int k = 0; k < rooms.length(); k++) {
									JSONObject room = rooms.getJSONObject(k);
									p1 = row.getCell(0).getParagraphs().get(0);
									p1.setAlignment(ParagraphAlignment.LEFT);
									XWPFRun r1 = p1.createRun();
									r1.setFontSize(8);
									r1.setText(room.toString());
									r1.addBreak();
								}
							}

//							row = tab.createRow(); // Second Row
//							row.getCell(0).setText("");
//							if (row.getCell(0).getCTTc().getTcPr() == null)
//								row.getCell(0).getCTTc().addNewTcPr();
//							if (row.getCell(0).getCTTc().getTcPr().getGridSpan() == null)
//								row.getCell(0).getCTTc().getTcPr().addNewGridSpan();
//							row.getCell(0).getCTTc().getTcPr().getGridSpan().setVal(BigInteger.valueOf((long) 3));
//
//							if (listing_history != null) {
//								for (int k = 0; k < listing_history.length(); k++) {
//									JSONObject listing = listing_history.getJSONObject(k);
//									p1 = row.getCell(0).getParagraphs().get(0);
//									p1.setAlignment(ParagraphAlignment.LEFT);
//									XWPFRun r1 = p1.createRun();
//									r1.setFontSize(8);
//									r1.setText(listing.toString());
//									r1.addBreak();
//								}
//							}
//							

//							
//							
//
//							row = tab.createRow(); // Third Row
//							row.getCell(0).setText("2.");
//							//row.getCell(1).setText("Mohan");
//							//row.getCell(2).setText("mohan@gmail.com");

							document.write(out);
						} catch (Exception es) {
							System.out.println(es.getMessage());
						}
					}
				} catch (Exception e) {
					System.out.println(e.getMessage());
				}

			}
		}

	}

	private static void printMLSDetail(String mlsid) {

	}

	private static void spanCellsAcrossRow(XWPFTable table, int rowNum, int colNum, int span, String content) {
		XWPFTableCell cell = table.getRow(rowNum).getCell(colNum);
		cell.getCTTc().getTcPr().addNewGridSpan();
		cell.getCTTc().getTcPr().getGridSpan().setVal(BigInteger.valueOf((long) span));
		cell.setText(content);
	}

	private static void writeHouseDetail(int j, JSONObject jobj, boolean addHeader) {

		int k = 0;
		for (Iterator<Entry<String, String>> iterator = Heads.entrySet().iterator(); iterator.hasNext();) {
			Entry<String, String> setting = iterator.next();
			String key = setting.getKey();
			String val = setting.getValue();
			if (addHeader) {
				if (key.contains("@")) {
					try {
						String[] arrKeys = key.split("@", 2);
						System.out.println("test - " + arrKeys);
						houseD[j][k] = arrKeys[1];
					} catch (JSONException e) {
						e.printStackTrace();
					}
				} else {
					houseD[j][k] = key;
				}
			} else {
				if (key.contains("@")) {
					try {
						String[] arrKeys = key.split("@", 2);
						houseD[j][k] = val.equalsIgnoreCase("String")
								? jobj.getJSONObject(arrKeys[0]).getString(arrKeys[1])
								: jobj.getJSONObject(arrKeys[0]).getInt(arrKeys[1]);
					} catch (JSONException e) {
						e.printStackTrace();
					}
				} else {
					try {
						if (key.equalsIgnoreCase("id_listing")) {
							houseD[j][k] = "https://housesigma.com/web/en/house/" + jobj.getString(key);
						} else {
							houseD[j][k] = val.equalsIgnoreCase("String") ? jobj.getString(key) : jobj.getInt(key);
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