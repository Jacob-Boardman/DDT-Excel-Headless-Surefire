import static org.junit.Assert.*;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Arrays;
import java.util.Collection;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.After;
import org.junit.Before;
import org.junit.Test;
import org.junit.runner.RunWith;
import org.junit.runners.Parameterized;
import org.junit.runners.Parameterized.Parameters;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.phantomjs.PhantomJSDriver;
import org.openqa.selenium.support.PageFactory;

@RunWith(Parameterized.class)
public class ExcelParaTest {

	WebDriver driver;

	@Before
	public void setup() {
		System.setProperty("phantomjs.binary.path", Constants.PHANTOMPATH);
		driver = new PhantomJSDriver();
	}

	@After
	public void teardown() {
		driver.quit();
	}

	@Parameters
	public static Collection<Object[]> data() throws IOException {
		FileInputStream file = new FileInputStream(Constants.excelpath);
		XSSFWorkbook workbook = new XSSFWorkbook(file);
		XSSFSheet sheet = workbook.getSheetAt(0);

		Object[][] ob = new Object[sheet.getPhysicalNumberOfRows() - 1][4];

		// for (int rowNum = 1; rowNum < sheet.getPhysicalNumberOfRows(); rowNum++) {
		// ob[rowNum-1][0] = sheet.getRow(rowNum).getCell(0).getStringCellValue();
		// ob[rowNum-1][1] = sheet.getRow(rowNum).getCell(1).getStringCellValue();
		// ob[rowNum-1][2] = sheet.getRow(rowNum).getCell(2).getStringCellValue();
		// ob[rowNum-1][3] = rowNum;
		// }

		for (int rowNum = 1; rowNum < sheet.getPhysicalNumberOfRows(); rowNum++) {
			for (int colNum = 0; colNum < 3; colNum++) {
				ob[rowNum - 1][colNum] = sheet.getRow(rowNum).getCell(colNum).getStringCellValue();
			}
			ob[rowNum - 1][3] = rowNum;
		}

		file.close();
		return Arrays.asList(ob);
	}

	private String username;
	private String password;
	private String expected;
	private int row;

	public ExcelParaTest(String username, String password, String expected, int row) {
		this.username = username;
		this.password = password;
		this.expected = expected;
		this.row = row;
	}

	@Test
	public void login() throws IOException {

		System.out.println(username + " " + password + " " + expected + " " + row);

		driver.manage().window().maximize();
		driver.get("http://thedemosite.co.uk/addauser.php");

		ExcelDemoLandingPage edlp = PageFactory.initElements(driver, ExcelDemoLandingPage.class);

		edlp.createAndLogin(username, password);
		WebElement login = driver.findElement(By.xpath(Constants.login));
		login.click();
		edlp.createAndLogin(username, password);

		FileInputStream file = new FileInputStream(Constants.excelpath);
		XSSFWorkbook workbook = new XSSFWorkbook(file);
		XSSFSheet sheet = workbook.getSheetAt(0);

		XSSFRow rowy = sheet.getRow(row);
		XSSFCell cellActual = rowy.getCell(3);
		if (cellActual == null) {
			cellActual = rowy.createCell(3);
		}

		WebElement text = driver.findElement(By.xpath(Constants.text));
		cellActual.setCellValue(text.getText());

		XSSFCell cellResult = rowy.getCell(4);
		if (cellResult == null) {
			cellResult = rowy.createCell(4);
		}
		try {
			assertEquals("Login not successful", expected, text.getText());

			cellResult.setCellValue("Pass");

		} catch (AssertionError e) {

			cellResult.setCellValue("Fail");
		} finally {

			FileOutputStream fileOut = new FileOutputStream(Constants.excelpath);

			workbook.write(fileOut);
			fileOut.flush();
			file.close();
		}
	}

}
