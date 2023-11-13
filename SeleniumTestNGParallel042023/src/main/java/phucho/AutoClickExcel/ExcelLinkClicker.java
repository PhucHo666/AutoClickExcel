package phucho.AutoClickExcel;

import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.openqa.selenium.By;
import org.openqa.selenium.TimeoutException;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;

import java.awt.Color;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Map;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;
import java.util.concurrent.TimeUnit;
import java.time.Duration;

public class ExcelLinkClicker {
    public static void main(String[] args) {
        // Set the path to your Excel file
        String excelFilePath = "src/main/resources/Index ClassPad Learning.xlsx";

        // Set up ChromeDriver
        System.setProperty("webdriver.chrome.driver", "/Users/phucho/Downloads/chromedriver-mac-x64/chromedriver");
        ChromeOptions options = new ChromeOptions();
        //options.addArguments("--start-maximized"); // Optionally maximize the browser window
        WebDriver driver = new ChromeDriver(options);

        try {
            // Load the Excel file
            FileInputStream excelFile = new FileInputStream(new File(excelFilePath));
            Workbook workbook = new XSSFWorkbook(excelFile);
            Sheet sheet = workbook.getSheetAt(0);
            ArrayList<String> errorList = new ArrayList<String>();

            int successCount = 0; // Initialize the success count


            for (Row row : sheet) {
                Cell cell = row.getCell(0); // Assuming links are in the first column (column 0)

                if (cell != null) {
                    Hyperlink hyperLink = cell.getHyperlink();
                    if (hyperLink == null) {
                        continue;
                    }
                    String link = hyperLink.getAddress(); // Trim any extra whitespace

                    if (link != null && !link.isEmpty()) {
                        //System.out.println("Opening link: " + link + " successfully"); // Debug statement
                        try {
                            // Open the link in the browser
                            driver.get(link);

                        } catch (Exception e) {
                            System.out.println("Error opening link: " + e.getMessage());
                        }
                        //Use WebDriverWait to wait for a specific condition, e.g., the presence of a certain element
                        Duration timeoutDuration = Duration.ofSeconds(10);
                        WebDriverWait wait = new WebDriverWait(driver, timeoutDuration); // Wait for up to 10 seconds

                        // Example: Wait for an element with a specific ID to be present
                        By elementLocator = By.xpath("//h1[normalize-space()='Login']");
                        // Wait for the element to be present
                        try {
                            WebElement element = wait.until(ExpectedConditions.presenceOfElementLocated(elementLocator));
                            successCount++; // Increment the success count
                            java.util.Date date = new java.util.Date();
                            System.out.println("(" + successCount + ")" + " Opening link: " + link + " successfully at " + date);
                            ExcelHelpers excel = new ExcelHelpers();
                            excel.setExcelFile(excelFilePath, "Sheet1");
                            excel.setCellData("Pass", cell.getRowIndex(), cell.getColumnIndex() + 1);

                        } catch (TimeoutException e) {
                            successCount++; // Increment the success count
                            java.util.Date date = new java.util.Date();
                            System.out.println("(" + successCount + ")" +" Error opening link: " + link + " failed at " + date);
                            ExcelHelpers excel = new ExcelHelpers();
                            excel.setExcelFile(excelFilePath, "Sheet1");
                            excel.setCellData("Fail", cell.getRowIndex(), cell.getColumnIndex() + 1);
                            excel.setCellData(link, cell.getRowIndex(), cell.getColumnIndex() +2);
                        }

                    }
                }
            }
            // Close the browser
            driver.quit();
            if (errorList.size() > 0) {
                System.out.println("Error opening: " + errorList);
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}

