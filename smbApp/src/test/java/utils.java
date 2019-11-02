
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.PreparedStatement;
import java.text.DateFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.Iterator;
import java.util.concurrent.TimeUnit;

import org.apache.commons.io.FileUtils;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.testng.annotations.Test;

public class utils {
	public FileInputStream fis = null;
	public XSSFWorkbook workbook = null;
	public XSSFSheet sheet = null;
	public XSSFRow row = null;
	public XSSFCell cell = null;
	String outputfilelocation;
	String hour;
	String mins;
	String secs;
	String time = java.time.LocalDateTime.now().toString();
	static String bcrmurl = "https://bcrm365.etisalat.corp.ae/Etisalat365/main.aspx";
	
	public String File_Name = "D:\\Spotfire\\Orders.xlsx";
	
	 TakesScreenshot driver;
	private static String[] columns = {"Orderlineitem", "Status", "Time", };

	public utils() {
		// TODO Auto-generated constructor stub
	}

	@Test
	public void testdate() throws ParseException {
		// CreationTime 11/02/2019 15:14
		// Completion Time 11/02/2019 15:21
		// Estimated Closure Date 21-Oct-19
		SimpleDateFormat formatter = new SimpleDateFormat("dd-MMM-yy");
		Date date = formatter.parse("21-Oct-19");
		// String date = "2019-08-24";
		String strDate = formatter.format(date);
		System.out.println("Date Format with yyyy-mm-dd : " + strDate);

		formatter = new SimpleDateFormat("dd-M-yyyy hh:mm:ss");
		strDate = formatter.format(date);
		System.out.println("Date Format with dd-M-yyyy hh:mm:ss : " + strDate);

		formatter = new SimpleDateFormat("yyyy-MM-dd");
		strDate = formatter.format(date);
		System.out.println("Date Format with dd MMMM yyyy : " + strDate);

		formatter = new SimpleDateFormat("dd MMMM yyyy zzzz");
		strDate = formatter.format(date);
		System.out.println("Date Format with dd MMMM yyyy zzzz : " + strDate);

		formatter = new SimpleDateFormat("E, dd MMM yyyy HH:mm:ss z");
		strDate = formatter.format(date);
		System.out.println("Date Format with E, dd MMM yyyy HH:mm:ss z : " + strDate);
	}

	// order_detail_title
public static void clickOnOrderDetailTab(WebDriver driver) {
		driver.findElement(By.id("order_detail_title")).click();
	}

	public static void clickOnOrderTab(WebDriver driver) {
		driver.findElement(By.id("order_title")).click();
		// order_title
	}

	public static void clickOnSubrequestDetail(WebDriver driver) {
		driver.findElement(By.id("sub_request_title")).click();

	}

public static WebDriver launchBrowser() {
		System.setProperty("webdriver.chrome.driver", "Z:\\Projects\\FelixData\\Production\\76\\chromedriver.exe");
	//	System.setProperty("webdriver.chrome.driver", "X:\\Automation\\Selenium\\76\\chromedriver.exe");
		ChromeOptions options = new ChromeOptions();
		options.addArguments("--allow-running-insecure-content");
	//	options.addArguments("headless");
		WebDriver driver = new ChromeDriver(options);
		driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
		driver.manage().window().fullscreen();
		return driver;
	}

	public static void readExcel(String fileName1) {
		// Creating a Workbook from an Excel file (.xls or .xlsx)
		// Connection con = connect();
		String line = "";
		PreparedStatement pstmt = null;
		System.out.println("Start time" + System.currentTimeMillis());
		String sql = "INSERT INTO BCRMTEST(OrderLineReferenceNumber,\r\n" + "OrderID,\r\n" + "CreatedOn,\r\n"
				+ "ClosedDate,\r\n" + "RequestStatus" + ") VALUES (?,?,?,?,?)";

		try {
			Workbook workbook = WorkbookFactory.create(new File(fileName1));
			// Retrieving the number of sheets in the Workbook
			System.out.println("Workbook has " + workbook.getNumberOfSheets() + " Sheets  ");

			Iterator<Sheet> sheetIterator = workbook.sheetIterator();
			System.out.println("Retrieving Sheets using Iterator");
			while (sheetIterator.hasNext()) {
				Sheet sheet = sheetIterator.next();
				System.out.println("=> " + sheet.getSheetName());
			}

			// Getting the Sheet at index zero

			Sheet sheet = workbook.getSheetAt(0);

			// Create a DataFormatter to format and get each cell's value as String
			DataFormatter dataFormatter = new DataFormatter();

			// 1. You can obtain a rowIterator and columnIterator and iterate over them
			System.out.println("\n\nIterating over Rows and Columns using Iterator\n");
			Iterator<Row> rowIterator = sheet.rowIterator();
			while (rowIterator.hasNext()) {
				Row row = rowIterator.next();
				System.out.println("Number of columns" + row.getLastCellNum());

				int lastColumn = Math.max(row.getLastCellNum(), 5);
				// Now let's iterate over the columns of the current row
				Iterator<Cell> cellIterator = row.cellIterator();
				int i = 1;
				for (int cn = 0; cn < lastColumn; cn++) {
					@SuppressWarnings("deprecation")
					Cell c = row.getCell(cn);
					String cellValue = dataFormatter.formatCellValue(c);
					if (cellValue == "") {
						cellValue = null;
					}
					System.out.println("Cell Value at :" + i + "is" + cellValue);
					i++;
					// cellValue.
				}

				System.out.println();
			}
			workbook.close();

		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}

	private static void writeExcel() {
		Workbook workbook1 = new XSSFWorkbook();
		CreationHelper createHelper = workbook1.getCreationHelper();
		// Create sheet
		Sheet sheet = workbook1.createSheet("BCRM");

		// Create row
		Row headerRow = sheet.createRow(0);
		int columnlength = 5;
		// Create cells
		for (int i = 0; i < columnlength; i++) {
			Cell cell = headerRow.createCell(i);
			cell.setCellValue("Test");
		}
	}

	// public ExcelApiTest(String xlFilePath) throws Exception
	// {
	// fis = new FileInputStream(xlFilePath);
	// workbook = new XSSFWorkbook(fis);
	// fis.close();
	// }
	String celvalue1 = "";

	@SuppressWarnings("resource")
	public ArrayList<String> excelRead(String File_Name, int sheetNumber, int columnNumber) throws FileNotFoundException {
		FileInputStream excelFile;
		String celvalue1;
		try {
			excelFile = new FileInputStream(new File(File_Name));
			Workbook workbook = new XSSFWorkbook(excelFile);
			Sheet datatypeSheet = workbook.getSheetAt(sheetNumber);
			Iterator<Row> iterator = datatypeSheet.iterator();
			ArrayList<String> arr =  new ArrayList<String>();;
			while (iterator.hasNext()) {
				Row currentRow = iterator.next();
				if (currentRow.getRowNum() == 0) {
					continue;
				}
				Cell cc = currentRow.getCell(columnNumber);
				System.out.println(cc.getCellType());
				CellType celltype = cc.getCellType();
				if (celltype.toString().contentEquals("STRING")) {
					celvalue1 = cc.getStringCellValue();
					System.out.println("String Cell value :" + celvalue1);
					
				}

				else if (celltype.toString().contentEquals("NUMERIC")) {
					double celvalue = cc.getNumericCellValue();
					System.out.println(" Numeric Cell value :" + celvalue);
					celvalue1 = Double.toString(celvalue);
				} else {
					Date celvalue = cc.getDateCellValue();
					celvalue1 = celvalue.toString();
					System.out.println("Date Cell value  :" + celvalue);
				}
				System.out.println("value " +celvalue1);
				
				arr.add(celvalue1);
				System.out.println("Array from utils : " + arr);
				
				
				
			}
			return arr;
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		return null;

	}
	public void excelRead(WebDriver driver ,String File_Name, int sheetNumber, int columnNumber) throws Exception {
		FileInputStream excelFile;
		String celvalue1;
		FileOutputStream fileOut = null;
		
		Workbook workbook1 = null;
		int wRowNum = 1;
		int count = 0;
			System.out.println("File name " + File_Name);
			excelFile = new FileInputStream(new File(File_Name));
			Workbook workbook = null;
			try {
				workbook = new XSSFWorkbook(excelFile);
			} catch (IOException e1) {
				// TODO Auto-generated catch block
			//Log.error("Eror message :"+	e1.getMessage());
			}
			
			try {
			Sheet datatypeSheet = workbook.getSheetAt(sheetNumber);
			Iterator<Row> iterator = datatypeSheet.iterator();
			workbook1 = new XSSFWorkbook(); // new HSSFWorkbook() for generating `.xls` file

			System.out.println("******* creation of sheet with name BCRM_ORDERS ******");
	        Sheet sheet1 = workbook1.createSheet("BCRM_ORDERS");
	        Font headerFont = workbook1.createFont();
	        headerFont.setBold(true);
	        headerFont.setFontHeightInPoints((short) 14);
	        headerFont.setColor(IndexedColors.BLACK.getIndex());
	        CellStyle headerCellStyle = workbook1.createCellStyle();
	        headerCellStyle.setFont(headerFont);
	        Row headerRow = sheet1.createRow(0);
	        for(int i = 0; i < columns.length; i++) {
	            Cell cell = headerRow.createCell(i);
	            cell.setCellValue(columns[i]);
	            cell.setCellStyle(headerCellStyle);
	        }
			while (iterator.hasNext()) {
				Row currentRow = iterator.next();
				if (currentRow.getRowNum() == 0) {
					continue;
				}
				Cell cc = currentRow.getCell(columnNumber);
				System.out.println(cc.getCellType());
				CellType celltype = cc.getCellType();
				if (celltype.toString().contentEquals("STRING")) {
					celvalue1 = cc.getStringCellValue();
					System.out.println("String Cell value :" + celvalue1);
					
				}

				else if (celltype.toString().contentEquals("NUMERIC")) {
					double celvalue = cc.getNumericCellValue();
					System.out.println(" Numeric Cell value :" + celvalue);
					celvalue1 = Double.toString(celvalue);
				} else {
					Date celvalue = cc.getDateCellValue();
					celvalue1 = celvalue.toString();
					System.out.println("Date Cell value  :" + celvalue);
				}
				System.out.println("value " +celvalue1);
				//*[@id="container"]/div/div[1]/div[1]/div[6]/input
				driver.findElement(By.className("text_search")).clear();
				//String order = foreach.next();
				driver.findElement(By.className("text_search")).sendKeys(celvalue1);// implement reading data from
				//Thread.sleep(2000);
				
				System.out.println(" Order line number " + ++count); // excel
				driver.findElement(By.className("search_btn")).click();
			// out file creation 
				System.out.println("Writing excel in progress");
			
				Row row1 = sheet1.createRow(wRowNum++);
				
				try {
					Thread.sleep(2000);
					String customerref = driver.findElement(By.xpath("//*[@id=\"orderList\"]/tbody/tr/td[12]/span")).getText();
					System.out.println("customerref :" + customerref);
					System.out.println("Order available in spotfire " + celvalue1);
					row1.createCell(0).setCellValue(celvalue1);
					row1.createCell(1).setCellValue("Available");
					row1.createCell(2).setCellValue(java.time.LocalDateTime.now().toString().toString());
				} catch (Exception e) {
					Thread.sleep(2000);
					String eNotfound = driver.findElement(By.xpath("//*[@id=\"container\"]/div/div[4]/p[1]")).getText();
					System.out.println("Order available in spotfire : "+celvalue1+" " + eNotfound);
					row1.createCell(0).setCellValue(celvalue1);
					row1.createCell(1).setCellValue("Not Available");
					row1.createCell(2).setCellValue(java.time.LocalDateTime.now().toString().toString());
					
				}

			}
			
			}
			catch(Exception e) {
			//	System.out.println("test");
				takeScreenshot(driver);
				System.out.println("Screenshot captured ");
			}
			
			finally
	        { 
				String username = System.getProperty("user.name");
				System.out.println(username);
				//File file = new file();
				time = java.time.LocalDateTime.now().toString();
				System.out.println("Current Time" + time);
				time.substring(0, 19);
				 hour = time.substring(11, 13).toString();
				 mins = time.substring(14, 16).toString();
				 secs = time.substring(17, 19).toString();
				System.out.println("Current Time" + time + " "+hour+ " "+mins+ " " + secs);
				outputfilelocation = "C:\\Users\\"+System.getProperty("user.name")+"\\Desktop\\Missing-orders-spotfire-"+ time.substring(0, 10)+"_"+hour+"_"+mins+"_"+secs+".xlsx";
				fileOut = new FileOutputStream(outputfilelocation);
				
				workbook1.write(fileOut);
				System.out.println("Analysis is avaiable at location : " + outputfilelocation); 
	        	fileOut.close();
	        	workbook1.close();
	        } 
	          
			}
		public void bcrmExcelRead(WebDriver driver, String URL, String File_Name, int sheetNumber, int columnNumber) throws Exception {
			FileInputStream excelFile;
			String celvalue1;
			FileOutputStream fileOut = null;
			
			Workbook workbook1 = null;
			int wRowNum = 1;
			int count = 0;
				System.out.println("File name " + File_Name);
				excelFile = new FileInputStream(new File(File_Name));
				Workbook workbook = null;
				try {
					workbook = new XSSFWorkbook(excelFile);
				} catch (IOException e1) {
					// TODO Auto-generated catch block
				//Log.error("Eror message :"+	e1.getMessage());
				}
				
				try {
				Sheet datatypeSheet = workbook.getSheetAt(sheetNumber);
				Iterator<Row> iterator = datatypeSheet.iterator();
				workbook1 = new XSSFWorkbook(); // new HSSFWorkbook() for generating `.xls` file

				System.out.println("******* creation sheet with name BCRM_ORDERS ******");
		        Sheet sheet1 = workbook1.createSheet("BCRM_ORDERS");
		        Font headerFont = workbook1.createFont();
		        headerFont.setBold(true);
		        headerFont.setFontHeightInPoints((short) 14);
		        headerFont.setColor(IndexedColors.BLACK.getIndex());
		        CellStyle headerCellStyle = workbook1.createCellStyle();
		        headerCellStyle.setFont(headerFont);
		        Row headerRow = sheet1.createRow(0);
		        for(int i = 0; i < columns.length; i++) {
		            Cell cell = headerRow.createCell(i);
		            cell.setCellValue(columns[i]);
		            cell.setCellStyle(headerCellStyle);
		        }
				while (iterator.hasNext()) {
					Row currentRow = iterator.next();
					if (currentRow.getRowNum() == 0) {
						continue;
					}
					Cell cc = currentRow.getCell(columnNumber);
					System.out.println(cc.getCellType());
					CellType celltype = cc.getCellType();
					if (celltype.toString().contentEquals("STRING")) {
						celvalue1 = cc.getStringCellValue();
						System.out.println("String Cell value :" + celvalue1);
						
					}

					else if (celltype.toString().contentEquals("NUMERIC")) {
						double celvalue = cc.getNumericCellValue();
						System.out.println(" Numeric Cell value :" + celvalue);
						celvalue1 = Double.toString(celvalue);
					} else {
						Date celvalue = cc.getDateCellValue();
						celvalue1 = celvalue.toString();
						System.out.println("Date Cell value  :" + celvalue);
					}
					System.out.println("value " +celvalue1);
					//*[@id="container"]/div/div[1]/div[1]/div[6]/input
					driver.get(URL);
					//----
					driver.findElement(By.className("text_search")).clear();
					//String order = foreach.next();
					driver.findElement(By.className("text_search")).sendKeys(celvalue1);// implement reading data from
					Thread.sleep(2000);
					
					System.out.println(" Order line number " + ++count); // excel
					driver.findElement(By.className("search_btn")).click();
				// out file creation 
					System.out.println("Writing excel in progress");
				
					Row row1 = sheet1.createRow(wRowNum++);
					
					try {
						Thread.sleep(2000);
						//reading the status of orderline items - get the list of entires and match the statu
						String customerref = driver.findElement(By.xpath("//*[@id=\"orderList\"]/tbody/tr/td[12]/span")).getText();
						System.out.println("customerref :" + customerref);
						System.out.println("Order available in spotfire " + celvalue1);
						row1.createCell(0).setCellValue(celvalue1);
						row1.createCell(1).setCellValue("Available");
						row1.createCell(2).setCellValue(java.time.LocalDateTime.now().toString().toString());
					} catch (Exception e) {
						Thread.sleep(2000);
						System.out.println(e);
						String eNotfound = driver.findElement(By.xpath("//*[@id=\"container\"]/div/div[4]/p[1]")).getText();
						System.out.println("Order available in spotfire : "+celvalue1+" " + eNotfound);
						row1.createCell(0).setCellValue(celvalue1);
						row1.createCell(1).setCellValue("Not Available");
						row1.createCell(2).setCellValue(java.time.LocalDateTime.now().toString().toString());
						
					}

				}
				
				}
				catch(Exception e) {
					System.out.println("test" +e);
					takeScreenshot(driver);
					System.out.println("Screenshot captured ");
				}
				
				finally
		        { 
					String username = System.getProperty("user.name");
					System.out.println(username);
					//File file = new file();
					time = java.time.LocalDateTime.now().toString();
					System.out.println("Current Time" + time);
					time.substring(0, 19);
					 hour = time.substring(11, 13).toString();
					 mins = time.substring(14, 16).toString();
					 secs = time.substring(17, 19).toString();
					System.out.println("Current Time" + time + " "+hour+ " "+mins+ " " + secs);
					outputfilelocation = "C:\\Users\\"+System.getProperty("user.name")+"\\Desktop\\Missing-orders-spotfire-"+ time.substring(0, 10)+"_"+hour+"_"+mins+"_"+secs+".xlsx";
					fileOut = new FileOutputStream(outputfilelocation);
					
					workbook1.write(fileOut);
					System.out.println("Analysis is avaiable at location : " + outputfilelocation); 
		        	fileOut.close();
		        	workbook1.close();
		        } 
		          
				
	}
	public void excelReadforcancelled(WebDriver driver,String File_Name, int sheetNumber, int columnNumber) throws Exception {
		FileInputStream excelFile;
		String celvalue1;
		FileOutputStream fileOut = null;
		driver.manage().timeouts().implicitlyWait(30, TimeUnit.SECONDS);
		System.out.println(" file name  " + File_Name);
		System.out.println("Column number " + columnNumber);
		Workbook workbook1 = null;
		int wRowNum = 1;
		int count = 0;
			System.out.println("File name " + File_Name);
			excelFile = new FileInputStream(new File(File_Name));
			Workbook workbook = null;
			try {
				workbook = new XSSFWorkbook(excelFile);
			} catch (IOException e1) {
				// TODO Auto-generated catch block
			//Log.error("Eror message :"+	e1.getMessage());
			}
			
			try {
			Sheet datatypeSheet = workbook.getSheetAt(sheetNumber);
			Iterator<Row> iterator = datatypeSheet.iterator();
			workbook1 = new XSSFWorkbook(); // new HSSFWorkbook() for generating `.xls` file

			System.out.println("******* creation sheet with name BCRM_ORDERS ******");
	        Sheet sheet1 = workbook1.createSheet("BCRM_ORDERS");
	        Font headerFont = workbook1.createFont();
	        headerFont.setBold(true);
	        headerFont.setFontHeightInPoints((short) 14);
	        headerFont.setColor(IndexedColors.BLACK.getIndex());
	        CellStyle headerCellStyle = workbook1.createCellStyle();
	        headerCellStyle.setFont(headerFont);
	        Row headerRow = sheet1.createRow(0);
	        for(int i = 0; i < columns.length; i++) {
	            Cell cell = headerRow.createCell(i);
	            cell.setCellValue(columns[i]);
	            cell.setCellStyle(headerCellStyle);
	        }
			while (iterator.hasNext()) {
				Row currentRow = iterator.next();
				if (currentRow.getRowNum() == 0) {
					continue;
				}
				Cell cc = currentRow.getCell(columnNumber);
				System.out.println(cc.getCellType());
				CellType celltype = cc.getCellType();
				if (celltype.toString().contentEquals("STRING")) {
					celvalue1 = cc.getStringCellValue();
					System.out.println("String Cell value :" + celvalue1);
					
				}

				else if (celltype.toString().contentEquals("NUMERIC")) {
					double celvalue = cc.getNumericCellValue();
					System.out.println(" Numeric Cell value :" + celvalue);
					celvalue1 = Double.toString(celvalue);
				} else {
					Date celvalue = cc.getDateCellValue();
					celvalue1 = celvalue.toString();
					System.out.println("Date Cell value  :" + celvalue);
				}
				System.out.println("value " +celvalue1);
				System.out.println(" order number : " +1);
				try {
				//*[@id="container"]/div/div[1]/div[1]/div[6]/input
					Thread.sleep(5000);
				driver.findElement(By.className("navImageFlipHorizontal")).click();// click on search button
				driver.findElement(By.cssSelector("#search")).sendKeys(celvalue1);
			//	 driver.findElement(By.cssSelector("#search")).sendKeys("ORD-1074591-H2K8F5");
				Thread.sleep(2000);
				 driver.findElement(By.id("findCriteriaButton")).click();// click on serach after entering the ord
				 Thread.sleep(2000);
				 WebElement allframe = driver.findElement(By.tagName("iframe"));
					System.out.println("Allframe :" + allframe);
				WebDriver testframe = driver.switchTo().frame(driver.findElement(By.tagName("iframe")));//bdb1bb2f9a43d1c23bba61c782ac6395
				 driver.findElement(By.xpath("//*[@id=\"attribone\"]")).click();
				//*[@id="customerid_lookupValue"]
				String customername =  driver.findElement(By.xpath("//*[@id=\"customerid_lookupValue\"]")).getText();
				System.out.println(" Customer name : " + customername);
				testframe.switchTo().defaultContent();
				if(customername.contains("UAT TEST")) {
					System.out.println(" cancellation in progress");
					driver.findElement(By.xpath("//*[@id=\"salesorder|NoRelationship|Form|Mscrm.Form.salesorder.CancelOrder\"]/span/a/span")).click();
				}
				else {
					System.out.println(" order cannot be cancelled");
				}
				
				 WebElement allframe1 = driver.findElement(By.tagName("iframe"));
					System.out.println("Allframe1 :" + allframe1);
					WebDriver frame = driver.switchTo().frame("InlineDialog_Iframe");
				driver.findElement(By.xpath("//*[@id=\"description\"]")).sendKeys("UAT TEST");
				driver.findElement(By.xpath("//*[@id=\"butBegin\"]")).click();
				driver.get(bcrmurl);
				
				///*[@id="description"]
				}
				catch(Exception e){
					System.out.println("Element Not found" +e);
					driver.get(bcrmurl);
				}
				
				
			// out file creation 
				System.out.println("Writing excel in progress");
			
			/*	Row row1 = sheet1.createRow(wRowNum++);
				row1.createCell(0).setCellValue(celvalue1);
				row1.createCell(1).setCellValue("Available");
				row1.createCell(2).setCellValue(java.time.LocalDateTime.now().toString().toString());
				
				try {
					Thread.sleep(2000);
					String customerref = driver.findElement(By.xpath("//*[@id=\"orderList\"]/tbody/tr/td[12]/span")).getText();
					System.out.println("customerref :" + customerref);
					System.out.println("Order available in spotfire " + celvalue1);
					row1.createCell(0).setCellValue(celvalue1);
					row1.createCell(1).setCellValue("Available");
					row1.createCell(2).setCellValue(java.time.LocalDateTime.now().toString().toString());
				} catch (Exception e) {
					Thread.sleep(2000);
					String eNotfound = driver.findElement(By.xpath("//*[@id=\"container\"]/div/div[4]/p[1]")).getText();
					System.out.println("Order available in spotfire : "+celvalue1+" " + eNotfound);
					row1.createCell(0).setCellValue(celvalue1);
					row1.createCell(1).setCellValue("Not Available");
					row1.createCell(2).setCellValue(java.time.LocalDateTime.now().toString().toString());
					
				}
*/
			}
			
			}
			catch(Exception e) {
				System.out.println("test");
				takeScreenshot(driver);
				System.out.println("Screenshot captured ");
			}
			
			finally
	        { 
				String username = System.getProperty("user.name");
				System.out.println(username);
				//File file = new file();
				time = java.time.LocalDateTime.now().toString();
				System.out.println("Current Time" + time);
				time.substring(0, 19);
				 hour = time.substring(11, 13).toString();
				 mins = time.substring(14, 16).toString();
				 secs = time.substring(17, 19).toString();
				System.out.println("Current Time" + time + " "+hour+ " "+mins+ " " + secs);
				outputfilelocation = "C:\\Users\\"+System.getProperty("user.name")+"\\Desktop\\Cancelled-orders-"+ time.substring(0, 10)+"_"+hour+"_"+mins+"_"+secs+".xlsx";
				fileOut = new FileOutputStream(outputfilelocation);
				
				workbook1.write(fileOut);
				System.out.println("Analysis is avaiable at location : " + outputfilelocation); 
	        	fileOut.close();
	        	workbook1.close();
	        } 
	          
			}
	
	public void takeScreenshot(WebDriver driver2)  {
		String time = java.time.LocalDateTime.now().toString();
		String hour = time.substring(11, 13).toString();
		String mins = time.substring(14, 16).toString();
		String secs = time.substring(17, 19).toString();
		File scrFile = ((TakesScreenshot)driver2).getScreenshotAs(OutputType.FILE);
        try {
			FileUtils.copyFile(scrFile, new File("C:\\Users\\" + System.getProperty("user.name") + "\\Desktop\\" +time.substring(0, 10)+"_"+hour+"_"+mins+"_"+secs+".png"));
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
    }
	public String getCellData(String sheetName, String colName, int rowNum) {
		try {
			int col_Num = -1;
			sheet = workbook.getSheet(sheetName);
			row = sheet.getRow(0);
			for (int i = 0; i < row.getLastCellNum(); i++) {
				if (row.getCell(i).getStringCellValue().trim().equals(colName.trim()))
					col_Num = i;
			}

			row = sheet.getRow(rowNum - 1);
			cell = row.getCell(col_Num);

			if (cell.getCellTypeEnum() == CellType.STRING)
				return cell.getStringCellValue();
			else if (cell.getCellTypeEnum() == CellType.NUMERIC || cell.getCellTypeEnum() == CellType.FORMULA) {
				String cellValue = String.valueOf(cell.getNumericCellValue());
				if (HSSFDateUtil.isCellDateFormatted(cell)) {
					DateFormat df = new SimpleDateFormat("dd/MM/yy");
					Date date = cell.getDateCellValue();
					cellValue = df.format(date);
				}
				return cellValue;
			} else if (cell.getCellTypeEnum() == CellType.BLANK)
				return "";
			else
				return String.valueOf(cell.getBooleanCellValue());
		} catch (Exception e) {
			e.printStackTrace();
			return "row " + rowNum + " or column " + colName + " does not exist  in Excel";
		}
	}
}
