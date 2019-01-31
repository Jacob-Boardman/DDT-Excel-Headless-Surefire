import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.FindBy;

public class ExcelDemoLandingPage {
	
	@FindBy(name = "username")
	private WebElement user;
	
	@FindBy(name = "password")
	private WebElement pass;
	
	@FindBy(name = "FormsButton2")
	private WebElement button;
	
	public void createAndLogin(String user, String pass) {
		this.user.sendKeys(user);
		this.pass.sendKeys(pass);
		button.click();
	}
}
