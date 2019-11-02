import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.concurrent.TimeUnit;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;

/**
 * 
 */

/**
 * @author arshakhan
 *
 */
public class BCRMordermappingComparision {

	static String spotfireUrl = "https://bamprod.etisalat.corp.ae/spotfire/wp/render/IVrckQRdr3gE4de7uB/analysis?file=/Etisalat/Reports/Project%20Felix%20Group/BusinessOrdersTracker&waid=GmJlXStVG0i3MNghvFa3l-211057d5bcE2sz&wavid=0";
	static String spotfireUrl1 = "https://bamprod.etisalat.corp.ae/spotfire/wp/LoggedOut.aspx?logoutByUser=True";
	static WebDriver driver = null;
	static String inputFilepath;
	static String columnNumber;

	private static String[] columns = { "Orderlineitem", "Status", "Time", };
	SimpleDateFormat date1 = new SimpleDateFormat("yyyy-MM-dd");

	public BCRMordermappingComparision() {
		// TODO Auto-generated constructor stub
	}

	/**
	 * @param args
	 * @throws Exception
	 * @throws NumberFormatException
	 */
	public static void main(String[] args) throws NumberFormatException, Exception {
		utils ut = new utils();
		int firstSheet = 0;
		int firstColumn = 3;

		try {
			System.out.println("*** Launch browser ****");
			driver = utils.launchBrowser();
			driver.manage().timeouts().implicitlyWait(30, TimeUnit.SECONDS);
			/// html/body/div[2]/div[3]/a
			System.out.println("*** open spotfire ****");
			driver.get(spotfireUrl);
			Thread.sleep(3000);
			// driver.findElement(By.xpath("/html/body/div[2]/div[3]/a")).click();
			// driver.findElement(By.xpath("//*[@id=\"library-navbar\"]/div[1]/div[1]/div[1]/library-sidebar-category/ul/li[1]/a")).click();
			// //*[@id="library-navbar"]/div[1]/div[2]/div[1]/div[2]/div[1]/div[2]/div[1]
			// driver.findElement(By.cssSelector("[title*='BusinessOrdersTracker']")).click();
			// Thread.sleep(3000);
			String pageTitle = driver.getTitle();
			System.out.println("Page Title :" + pageTitle);
			driver.manage().timeouts().implicitlyWait(30, TimeUnit.SECONDS);
			WebElement allframe = driver.findElement(By.tagName("iframe"));
			System.out.println("Allframe :" + allframe);
			driver.switchTo().frame(driver.findElement(By.tagName("iframe")));
			driver.manage().timeouts().implicitlyWait(30, TimeUnit.SECONDS);
			System.out.println("calling excel read");
			System.out.println("*** Excel file reading .... ****");
			
			if (args.length > 0) {
				inputFilepath = args[0];
				columnNumber = args[1];
				
				System.out.println("inputFilepath : " + inputFilepath);
				System.out.println("Column number : " + columnNumber);

				if (inputFilepath.length() > 5 && columnNumber.length() > 0) {
					System.out.println("*** File name passed from cmd *****");
					ut.excelRead(driver, inputFilepath + ".xlsx", firstSheet, Integer.valueOf(columnNumber) - 1);
				} else {
					System.out.println("*** Default excel file name  is BCRM *****");
					ut.excelRead(driver, "C:\\Users\\" + System.getProperty("user.name") + "\\Desktop\\BCRM.xlsx",
							firstSheet, firstColumn);
				}

			} else {
				System.out.println("*** Default excel file name  is BCRM *****");
				ut.excelRead(driver, "C:\\Users\\" + System.getProperty("user.name") + "\\Desktop\\BCRM.xlsx",
						firstSheet, firstColumn);
			}
		} finally {
			System.out.println("***** Closing the browser *************");
			driver.close();

		}
	}

	
}
