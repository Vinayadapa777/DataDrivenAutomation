package TestingIM;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.List;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;

public class Pdpautomation extends Resources {
    public static void main(String[] args) throws IOException, InterruptedException {
	WebDriver driver = new ChromeDriver();
	driver.manage().window().maximize();
	String path = System.getProperty("user.dir") + "\\InputFiles\\PDPTesting.xlsx";
	FileInputStream fis = new FileInputStream(path);
	try (XSSFWorkbook wb = new XSSFWorkbook(fis)) {
	    XSSFSheet sh = wb.getSheetAt(0);
	    for (int i = 1; i < sh.getLastRowNum(); i++) {
		String url = getDataOfColumn(i, "Url");
		driver.get(url);
		Thread.sleep(3000);
		try {
		    WebElement pdf = driver.findElement(By.linkText("Product Brochure"));
		    if (pdf.isDisplayed()) {
			setDataByColumnName1(i, "PDF", "Yes");
		    }
		} catch (Exception e) {
		    setDataByColumnName1(i, "No PDF", "Yes");
		}
		try {
		    WebElement singleImage = driver.findElement(By.cssSelector("#img_id"));
		    List<WebElement> multiImage = driver.findElements(By.xpath("//img[contains(@id,'mlt_img')]"));
		    int img_size = multiImage.size();
		    if (img_size > 1) {
			setDataByColumnName1(i, "Multi Image", "Yes");
		    } else if (img_size == 1 || singleImage.isDisplayed()) {
			setDataByColumnName1(i, "Single Image", "Yes");
		    }
		} catch (Exception e) {
		}
		try {
		    String video1="icon_video";
		    WebElement video = driver.findElement(By.xpath("//span[contains(@id,'"+video1+"')]"));
		    if (video.isDisplayed()) {
			setDataByColumnName1(i, "Video", "Yes");
		    }
		} catch (Exception e) {
		}
		Thread.sleep(3000);
		String con = getData(i + 1, 0);
		if (con == "" || con.isEmpty())
		    break;
	    }
	}
	driver.quit();
    }
}
