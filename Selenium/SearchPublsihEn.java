import java.io.File;
import java.io.FileNotFoundException;
import java.lang.reflect.Method;
import java.util.Scanner;

import jxl.Sheet;
import jxl.Workbook;

import org.junit.Test;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;

public class SearchPublsihEn {
	private WebDriver driver;
	private String baseUrl;

	@Test
	public void testBasfHomePage_Chrome() throws Exception {
		Scanner scanner = new Scanner(System.in);
		System.out
				.println("Please enter the excel file name without extension (Only xls format files allowed):");
		String fileName = scanner.next() + ".xls";
		scanner.close();
		Workbook workbook = null;
		try {
			workbook = Workbook.getWorkbook(new File(fileName));
		} catch (FileNotFoundException fileNotFoundException) {
			System.out.println("No excel file found with the specified name.");
			Thread.sleep(3000);
			System.exit(0);
		}

		System.setProperty("webdriver.chrome.driver",
				"C:\\Users\\dishant.a.chawla\\Downloads\\chromedriver.exe");

		// Sheet sheet = workbook.getSheet(0);
		for (Sheet sheet : workbook.getSheets()) {
			driver = new ChromeDriver();
			Actions action = new Actions(driver);

			WebDriverWait wait = new WebDriverWait(driver, 10);
			baseUrl = sheet.getCell(2, 1).getContents();
			driver.get(baseUrl);

			Method findElementMethod;
			WebElement element;
			String eventMethodName, parameters, findElementValue;
			int noOfRows = sheet.getRows();

			for (int r = 2; r < noOfRows; r++) {

				findElementMethod = By.class.getMethod(sheet.getCell(1, r)
						.getContents(), String.class);
				findElementValue = sheet.getCell(2, r).getContents();
				eventMethodName = sheet.getCell(3, r).getContents();
				parameters = sheet.getCell(4, r).getContents();

				if (sheet.getCell(0, r).getContents()
						.equalsIgnoreCase("Open URL")) {
					driver.navigate().to(findElementValue);
				}

				element = wait.until(ExpectedConditions
						.elementToBeClickable((By) findElementMethod.invoke(
								null, findElementValue)));
				if (eventMethodName.equalsIgnoreCase("sendKeys")) {
					element.getClass()
							.getMethod(eventMethodName, CharSequence[].class)
							.invoke(element,
									(Object) new String[] { parameters });
				} else if (eventMethodName.equalsIgnoreCase("contextClick")
						|| eventMethodName.equalsIgnoreCase("doubleClick")) {
					((Actions) action.getClass().getMethod(eventMethodName)
							.invoke(element)).perform();

				} else if (eventMethodName.equalsIgnoreCase("verify")) {
					if (element.getText().equals(parameters))
						System.out.println(parameters + " verified");
					else
						System.out.println(parameters + " verification failed");
				} else {
					element.getClass().getMethod(eventMethodName)
							.invoke(element);
				}
			}
		}
	}
}