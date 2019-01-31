
import static org.junit.Assert.assertEquals;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.util.*;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.After;
import org.junit.Before;
import org.junit.Test;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.phantomjs.PhantomJSDriver;
import org.openqa.selenium.support.PageFactory;

public class ExcelDemoTest {

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

	@Test
	public void test() throws Exception {
		driver.manage().window().maximize();
		driver.get("http://thedemosite.co.uk/addauser.php");
		ExcelDemoLandingPage edlp = PageFactory.initElements(driver, ExcelDemoLandingPage.class);

		FileInputStream file = new FileInputStream(Constants.excelpath);
		XSSFWorkbook workbook = new XSSFWorkbook(file);
		XSSFSheet sheet = workbook.getSheetAt(0);

		for (int rowNum = 1; rowNum < sheet.getPhysicalNumberOfRows(); rowNum++) {
			List<String> entries = new ArrayList<String>();
			for (int colNum = 0; colNum < sheet.getRow(rowNum).getPhysicalNumberOfCells(); colNum++) {
				XSSFCell cell = sheet.getRow(rowNum).getCell(colNum);
				String userCell = cell.getStringCellValue();
				entries.add(userCell);

			}
			System.out.println(entries);
			edlp.createAndLogin(entries.get(0), entries.get(1));
			WebElement login = driver.findElement(By.xpath(Constants.login));
			login.click();
			edlp.createAndLogin(entries.get(0), entries.get(1));
			
			XSSFRow row = sheet.getRow(rowNum);
			XSSFCell cellActual = row.getCell(3);
			XSSFCell cellResult = row.getCell(4);
			if (cellActual == null) {
				cellActual = row.createCell(3);
			}
			WebElement text = driver.findElement(By.xpath(Constants.text));
			cellActual.setCellValue(text.getText());
			
			if(cellResult == null) {
				cellResult = row.createCell(4);
			}
				
			if(entries.get(2).equals(text.getText())) {
				cellResult.setCellValue("Pass");								
			}
			else {
				cellResult.setCellValue("Fail");
			}
			
			
			WebElement adduser = driver.findElement(By.xpath(Constants.adduser));
			adduser.click();

		}
		
		FileOutputStream fileOut = new FileOutputStream(Constants.excelpath);
		
		workbook.write(fileOut);
		fileOut.flush();
		file.close();
	}

}
