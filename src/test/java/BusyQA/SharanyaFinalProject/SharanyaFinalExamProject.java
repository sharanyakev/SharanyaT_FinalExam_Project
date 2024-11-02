package BusyQA.SharanyaFinalProject;

import org.openqa.selenium.By;
import org.openqa.selenium.NoSuchElementException;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import io.github.bonigarcia.wdm.WebDriverManager;
import org.openqa.selenium.edge.EdgeDriver;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.OutputType;
import java.io.File;
import java.time.Duration;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;
import org.apache.commons.io.FileUtils;
import org.openqa.selenium.TakesScreenshot;
import com.aventstack.extentreports.ExtentReports;
import com.aventstack.extentreports.ExtentTest;
import com.aventstack.extentreports.reporter.ExtentSparkReporter;
import com.aventstack.extentreports.reporter.configuration.Theme;

public class SharanyaFinalExamProject {

    public static void main(String[] args) {

        ExtentSparkReporter sparkReporter = new ExtentSparkReporter("C:\\Users\\Mano\\eclipse-workspace\\SharanyaFinalProject\\src\\main\\resources\\extentReport.html");
        sparkReporter.config().setTheme(Theme.STANDARD);
        ExtentReports extent = new ExtentReports();
        extent.attachReporter(sparkReporter);

        WebDriverManager.edgedriver().setup();
        WebDriver driver = new EdgeDriver();

        Workbook workbook = new XSSFWorkbook();

        WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(50));

        try {
            driver.get("https://www.finmun.finances.gouv.qc.ca/finmun/f?p=100:3000::RESLT");

            wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("(//tbody[2]/tr[1]/td/a)[1]")));

            int count = 1;
            for (int i = 2; i < 6; i++) {
                for (int j = 1; j < 6; j++) {
                    try {
                        WebElement element = driver.findElement(By.xpath(String.format("(//tbody[%s]/tr[%s]/td/a)[1]", i, j)));
                        String sheet_name = element.getText();

                        File screenshot = ((TakesScreenshot) driver).getScreenshotAs(OutputType.FILE);
                        String screenshotPath = "C:\\Users\\Mano\\eclipse-workspace\\SharanyaFinalProject\\src\\main\\resources\\" + sheet_name + "screenshot1.png";
                        FileUtils.copyFile(screenshot, new File(screenshotPath));

                        ExtentTest test = extent.createTest("Test Case " + count++, sheet_name);
                        test.addScreenCaptureFromPath(screenshotPath);

                        element.click();

                        WebElement iframeElement = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//iframe[contains(@src, 'f?p=100:3015')]")));
                        driver.switchTo().frame(iframeElement);

                        wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//v[text()='SOUMISSION GAGNANTE']")));
                        WebElement table = driver.findElement(By.xpath("//p[v[text()='SOUMISSION GAGNANTE']]/following-sibling::table"));

                        File screenshot1 = ((TakesScreenshot) driver).getScreenshotAs(OutputType.FILE);
                        String screenshotPath1 = "C:\\Users\\Mano\\eclipse-workspace\\SharanyaFinalProject\\src\\main\\resources\\" + sheet_name + "screenshot2.png";
                        FileUtils.copyFile(screenshot1, new File(screenshotPath1));
                        test.addScreenCaptureFromPath(screenshotPath1);

                        List<WebElement> rows = table.findElements(By.tagName("tr"));
                        Sheet sheet = workbook.createSheet(sheet_name);
                        int rowNum = 0;
                        for (WebElement row : rows) {
                            Row excelRow = sheet.createRow(rowNum++);
                            List<WebElement> cells = row.findElements(By.tagName("td"));
                            int cellNum = 0;
                            for (WebElement cell : cells) {
                                Cell excelCell = excelRow.createCell(cellNum++);
                                excelCell.setCellValue(cell.getText());
                            }
                        }

                        test.pass("Table data extracted and screenshot attached.");

                        driver.switchTo().defaultContent();

                        WebElement close_button = driver.findElement(By.xpath("//button[text()='Fermer']"));
                        close_button.click();

                    } catch (NoSuchElementException e) {                    	System.out.println("No Such Element");
                        break;
                    } catch (Exception e) {
                    	 System.out.println("An error occurred");
                    }
                }
            }

            try (FileOutputStream fileOut = new FileOutputStream("C:\\Users\\Mano\\eclipse-workspace\\SharanyaFinalProject\\table_data.xlsx")) {
                workbook.write(fileOut);
            } catch (IOException e) {
                System.out.println("An error occurred");
            }

        } finally {
            // Close the workbook
            try {
                workbook.close();
            } catch (IOException e) {
            	System.out.println("An error occurred");
            }
            
            driver.quit();
            extent.flush();
        }

        System.out.println("Table data has been written to table_data.xlsx and extentReport.html generated.");
    }
}
