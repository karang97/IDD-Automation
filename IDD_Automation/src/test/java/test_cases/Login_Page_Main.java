package test_cases;
import java.io.BufferedReader;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.IOException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebElement;

import base.Base;

public class Login_Page_Main extends Base  {
	
	public static void Fap() throws InvalidFormatException, IOException, InterruptedException{
		driver=Base.openBrowser();
		driver.findElement(By.xpath("//span[contains(text(),'Insights')]")).click();
		driver.findElement(By.xpath("//span[text()=' Persona Components']")).click();
		driver.findElement(By.xpath("//*[text()='Total Number of Function Access Profiles']")).click();

		WebElement dropdown=driver.findElement(By.xpath("(//div[@class='dmm-dropdown-bottom dropdown'])[2]"));
		dropdown.click();
		driver.findElement(By.xpath("//div[normalize-space()='>']")).click();
		driver.findElement(By.xpath("//input[@type='number']")).sendKeys("0");
		driver.findElement(By.xpath("//input[@type='number']")).sendKeys(Keys.ENTER);
		driver.findElement(By.xpath("//button[text()=' Export ']")).click();

		System.out.println("---------------------------------------------");
		//System.out.println(driver.findElement(By.xpath("//div[@class='insights-detail-table']//descendant::div[@class='table-height ng-star-inserted']")).getText());
	}	

	public static void convertCSVToExcel(String csvFilePath, String excelFilePath) throws IOException {
		 // Create a new Workbook and Sheet
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("Sheet1");
        // Read the CSV file
        try (BufferedReader br = new BufferedReader(new FileReader(csvFilePath))) {
            String line;
            int rowIndex = 1;
            	while ((line = br.readLine()) != null) {
                // Split the line by comma (assuming comma-separated CSV)
                String[] cells = line.split(",");

                // Create a new row in the Excel sheet
                XSSFRow row = sheet.createRow(rowIndex++);
                // Trim spaces and write each cell
                for (int i = 0; i < cells.length; i++) {
                    // Trim leading/trailing spaces from each cell
                    String trimmedCell = cells[i].replaceAll("\\s{2,}", " ").replaceAll(",", "");
                    // Create a new cell in the row
                    XSSFCell cell = row.createCell(i);
                    cell.setCellValue(trimmedCell);

                }
            }
        }

        // Write the workbook to an XLSX file
        try (FileOutputStream fos = new FileOutputStream(excelFilePath)) {
            workbook.write(fos);
        }

        // Close the workbook
        workbook.close();
        
        System.out.println("Conversion completed: " + excelFilePath);
    }
	public static void ExcelCopy(){

		String sourceFilePath = "C:\\files\\SOURCE.xlsx";
		String destinationFilePath = "C:\\Users\\gadhavek\\OneDrive - Automatic Data Processing Inc\\Desktop\\NAV Tool - IDD Instructions and Client Approval v2.xlsx";

		try (FileInputStream sourceFile = new FileInputStream(sourceFilePath);
				FileInputStream destinationFile = new FileInputStream(destinationFilePath)) {

			// Load the source and destination workbooks
			XSSFWorkbook sourceWorkbook = new XSSFWorkbook(sourceFile);
			XSSFWorkbook destinationWorkbook = new XSSFWorkbook(destinationFile);

			// Get the specific sheet from both workbooks
			XSSFSheet sourceSheet = sourceWorkbook.getSheetAt(0); // Adjust sheet index as needed
			XSSFSheet destinationSheet = destinationWorkbook.getSheetAt(6); // Adjust sheet index as needed

			// Iterate through the rows and cells of the source sheet and copy them to the destination sheet
			int rowCount = sourceSheet.getLastRowNum();			

			for (int i = 1; i <= rowCount; i++) {
				Row sourceRow = sourceSheet.getRow(i);
				Row destinationRow = destinationSheet.createRow(i);

				int cellCount = sourceRow.getLastCellNum();
				for (int j = 0; j < cellCount; j++) {
					Cell sourceCell = sourceRow.getCell(j);
					Cell destinationCell = destinationRow.createCell(j);
					//cellValue = cellValue.replaceAll("\\s{2,}", " ").replaceAll(",", "");
                   // targetCell.setCellValue(cellValue);

					

					// Copy data based on cell type
					switch (sourceCell.getCellType()) {
					case STRING:
						destinationCell.setCellValue(sourceCell.getStringCellValue());
						break;
					case NUMERIC:
						destinationCell.setCellValue(sourceCell.getNumericCellValue());
						break;
					case BOOLEAN:
						destinationCell.setCellValue(sourceCell.getBooleanCellValue());
						break;
					case FORMULA:
						destinationCell.setCellFormula(sourceCell.getCellFormula());
						break;
					default:
						break;
					}
				}
			}


			// Write the changes to the destination workbook
			try (FileOutputStream outputStream = new FileOutputStream(destinationFilePath)) {
				destinationWorkbook.write(outputStream);
			}

			System.out.println("Data copied successfully!");

		} catch (IOException e) {
			e.printStackTrace();
		}
	}


	public static void main(String[] args) throws InvalidFormatException, InterruptedException, IOException {
		Fap();
		String csvFilePath = "C:\\Users\\gadhavek\\Downloads\\Persona Components - Function Access Profiles.csv";
		String excelFilePath = "C:\\files\\SOURCE.xlsx";
		convertCSVToExcel(csvFilePath, excelFilePath);
		ExcelCopy();
        FileInputStream fis = new FileInputStream("C:\\Users\\gadhavek\\OneDrive - Automatic Data Processing Inc\\Desktop\\NAV Tool - IDD Instructions and Client Approval v2.xlsx");
        XSSFWorkbook workbook = new XSSFWorkbook(fis);

        // Iterate over all sheets
        for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
            XSSFSheet sheet = workbook.getSheetAt(i);

            // Iterate over all rows in the sheet
            for (Row row : sheet) {
                // Iterate over all cells in the row
                for (Cell cell : row) {
                    if (cell.getCellType() == CellType.STRING) {
                        String cellValue = cell.getStringCellValue();
                        
                        // Replace double spaces with a single space
                        cellValue = cellValue.replaceAll("\\s{2,}", " ");
                        // Remove commas
                        cellValue = cellValue.replace(",", "");

                        // Set the updated value back to the cell
                        cell.setCellValue(cellValue);
                    }
                }
            }
        }

        // Close the input file stream
        fis.close();

        // Save the changes to a new file
        FileOutputStream fos = new FileOutputStream("C:\\Users\\gadhavek\\OneDrive - Automatic Data Processing Inc\\Desktop\\NAV Tool - IDD Instructions and Client Approval v2.xlsx");
        workbook.write(fos);
        fos.close();

        // Close the workbook
        workbook.close();

        System.out.println("Double spaces and commas removed from all sheets.");
    }

	}
