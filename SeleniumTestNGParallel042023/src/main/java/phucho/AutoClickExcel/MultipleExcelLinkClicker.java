package phucho.AutoClickExcel;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;

import java.io.File;
import java.io.FileInputStream;
import java.time.Duration;
import java.util.ArrayList;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;
import java.util.concurrent.TimeUnit;
import java.util.concurrent.atomic.AtomicInteger;

public class MultipleExcelLinkClicker {
    public static void main(String[] args) {
        // Set the path to your Excel file
        String excelFilePath = "src/main/resources/Index ClassPad Learning.xlsx";

        // Set up ChromeDriver
        System.setProperty("webdriver.chrome.driver", "/Users/phucho/Downloads/chromedriver-mac-x64/chromedriver");
        ChromeOptions options = new ChromeOptions();
        // options.addArguments("--start-maximized"); // Optionally maximize the browser window
        WebDriver driver = new ChromeDriver(options);

        try {
            // Load the Excel file
            FileInputStream excelFile = new FileInputStream(new File(excelFilePath));
            Workbook workbook = new XSSFWorkbook(excelFile);
            Sheet sheet = workbook.getSheetAt(0);
            ArrayList<String> errorList = new ArrayList<String>();

            AtomicInteger successCount = new AtomicInteger(); // Initialize the success count

            // Number of threads to run simultaneously
            int threadPoolSize = 3;
            ExecutorService executorService = Executors.newFixedThreadPool(threadPoolSize);

            for (Row row : sheet) {
                Cell cell = row.getCell(0); // Assuming links are in the first column (column 0)

                if (cell != null) {
                    Hyperlink hyperLink = cell.getHyperlink();
                    if (hyperLink == null) {
                        continue;
                    }
                    String link = hyperLink.getAddress(); // Trim any extra whitespace

                    if (link != null && !link.isEmpty()) {
                        executorService.submit(() -> {
                            WebDriver threadDriver = null;
                            try {
                                // Open the link in the browser
                                threadDriver = new ChromeDriver(options);
                                threadDriver.get(link);

                                // ... Perform actions on the opened link ...
                                // For example, wait for an element and print its text
                                By elementLocator = By.xpath("//h1[normalize-space()='Login']");
                                WebDriverWait wait = new WebDriverWait(threadDriver, Duration.ofSeconds(10));
                                WebElement element = wait.until(ExpectedConditions.presenceOfElementLocated(elementLocator));
                                successCount.getAndIncrement(); // Increment the success count
                                java.util.Date date = new java.util.Date();
                                System.out.println("(" + successCount + ") " + "Open: " + link + " successfully" + " at " + date);
                                ExcelHelpers excel = new ExcelHelpers();
                                excel.setExcelFile(excelFilePath, "Sheet1");
                                excel.setCellData("Pass", cell.getRowIndex(), cell.getColumnIndex() + 1);

                                // Close the browser
                                driver.quit();

                            } catch (Exception e) {
                                errorList.add(link);
                                successCount.getAndIncrement(); // Increment the success count
                                System.out.println("(" + successCount + ")" + "Error opening link: " + link + " - " + e.getMessage());
                                ExcelHelpers excel = new ExcelHelpers();

                                try {
                                    excel.setExcelFile(excelFilePath, "Sheet1");
                                } catch (Exception ex) {
                                    throw new RuntimeException(ex);
                                }
                                excel.setCellData("Fail", cell.getRowIndex(), cell.getColumnIndex() + 1);
                                excel.setCellData(link, cell.getRowIndex(), cell.getColumnIndex() + 2);
                                driver.quit();

                            }
                        });
                    }
                }
            }
            // Shut down the executor service gracefully
            executorService.shutdown();
            try {
                executorService.awaitTermination(Long.MAX_VALUE, TimeUnit.SECONDS);
            } catch (InterruptedException e) {
                e.printStackTrace();
            }

            if (errorList.size() > 0) {
                System.out.println("Error opening links: " + errorList);
            }
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            // Close the browser
            driver.quit();
        }
    }
}
