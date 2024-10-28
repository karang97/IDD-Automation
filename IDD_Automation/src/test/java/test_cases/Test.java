package test_cases;
import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.IOException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;

import base.*;
public class Test extends Base  {
	public static void Client_name() throws InvalidFormatException, IOException {
		driver.findElement(By.xpath("//a[normalize-space()='Projects']")).click();
		XSSFWorkbook Credentials = new XSSFWorkbook(new File("C:\\softwares and jars\\IDD_Automation\\test files\\login_details.xlsx"));
		XSSFSheet Sheet = Credentials.getSheetAt(0);
		String client =Sheet.getRow(1).getCell(4).toString().trim();
		WebElement element=driver.findElement(By.xpath("//span[text()='"+client+"']"));
		element.click();
	}
	
	
	public static void Fap() throws InvalidFormatException, IOException, InterruptedException{
		driver=Base.openBrowser();
		System.out.println("FAP START");
		String csvFilePath = "C:\\Users\\gadhavek\\Downloads\\Persona Components - Function Access Profiles.csv";
		String excelFilePath = "C:\\softwares and jars\\IDD_Automation\\test files\\SOURCE.xlsx";
		
		driver.findElement(By.xpath("//em[@class='icon icon-Insight']")).click();
		driver.findElement(By.xpath("//span[text()=' Persona Components']")).click();
		driver.findElement(By.xpath("//*[text()='Total Number of Function Access Profiles']")).click();

		WebElement dropdown=driver.findElement(By.xpath("(//div[@class='dmm-dropdown-bottom dropdown'])[2]"));
		dropdown.click();
		driver.findElement(By.xpath("//div[normalize-space()='>']")).click();
		driver.findElement(By.xpath("//input[@type='number']")).sendKeys("0");
		driver.findElement(By.xpath("//input[@type='number']")).sendKeys(Keys.ENTER);
		driver.findElement(By.xpath("//button[text()=' Export ']")).click();
		System.out.println("HALT FOR 1 SEC");
		Thread.sleep(1000);
		System.out.println("---------------------------------------------");
		XSSFWorkbook workbook = new XSSFWorkbook();
		XSSFSheet sheet = workbook.createSheet();
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
					String trimmedCell = cells[i].replaceAll("^\"|\"$", "");

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
		workbook.close();
		System.out.println("Conversion completed: " + excelFilePath);
		String sourceFilePath = "C:\\softwares and jars\\IDD_Automation\\test files\\SOURCE.xlsx";
		String destinationFilePath = "C:\\softwares and jars\\IDD_Automation\\test files\\NAV Tool - IDD Instructions and Client Approval v2.xlsx";
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
			System.out.println("FAP END");

		} catch (IOException e) {
			e.printStackTrace();
		}
	}


	//----------------------------------------------DISPLAY PROFILE----------------------------------------------------------

	public static void Display_profile() throws InvalidFormatException, InterruptedException, IOException {
		System.out.println("DISPLAY START");
		//driver=Base.openBrowser();
		String csvFilePath = "C:\\Users\\gadhavek\\Downloads\\Persona Components - Display Profiles.csv";
		String excelFilePath = "C:\\softwares and jars\\IDD_Automation\\test files\\SOURCE.xlsx";
		Client_name();
		
		driver.findElement(By.xpath("//em[@class='icon icon-Insight']")).click();
		driver.findElement(By.xpath("//span[text()=' Persona Components']")).click();
		driver.findElement(By.xpath("//*[text()='Total Number of Display Profiles']")).click();

		WebElement dropdown=driver.findElement(By.xpath("(//div[@class='dmm-dropdown-bottom dropdown'])[2]"));
		dropdown.click();
		driver.findElement(By.xpath("//div[normalize-space()='>']")).click();
		driver.findElement(By.xpath("//input[@type='number']")).sendKeys("0");
		driver.findElement(By.xpath("//input[@type='number']")).sendKeys(Keys.ENTER);
		driver.findElement(By.xpath("//button[text()=' Export ']")).click();
		System.out.println("HALT FOR 1 SEC");
		Thread.sleep(1000);		
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
					String trimmedCell = cells[i].replaceAll("^\"|\"$", "");
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
		
		workbook.close();
		System.out.println("Conversion completed: " + excelFilePath);
		String sourceFilePath = "C:\\softwares and jars\\IDD_Automation\\test files\\SOURCE.xlsx";
		String destinationFilePath = "C:\\softwares and jars\\IDD_Automation\\test files\\NAV Tool - IDD Instructions and Client Approval v2.xlsx";
		try (FileInputStream sourceFile = new FileInputStream(sourceFilePath);
				FileInputStream destinationFile = new FileInputStream(destinationFilePath)) {
			// Load the source and destination workbooks
			XSSFWorkbook sourceWorkbook = new XSSFWorkbook(sourceFile);
			XSSFWorkbook destinationWorkbook = new XSSFWorkbook(destinationFile);
			// Get the specific sheet from both workbooks
			XSSFSheet sourceSheet = sourceWorkbook.getSheetAt(0); // Adjust sheet index as needed
			XSSFSheet destinationSheet = destinationWorkbook.getSheetAt(7); // Adjust sheet index as needed
			// Iterate through the rows and cells of the source sheet and copy them to the destination sheet
			int rowCount = sourceSheet.getLastRowNum();
			for (int i = 1; i <= rowCount; i++) {
				Row sourceRow = sourceSheet.getRow(i);
				Row destinationRow = destinationSheet.createRow(i);

				int cellCount = sourceRow.getLastCellNum();
				for (int j = 0; j < cellCount; j++) {
					Cell sourceCell = sourceRow.getCell(j);
					Cell destinationCell = destinationRow.createCell(j);

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
			System.out.println("Display END");

		} catch (IOException e) {
			e.printStackTrace();
		}
	}
	//------------------------------------------------PAY RULES-----------------------------------------
	public static void Pay_Rules() throws InvalidFormatException, InterruptedException, IOException {
		//driver=Base.openBrowser();
		System.out.println("PAY_Rule START");
		String csvFilePath = "C:\\Users\\gadhavek\\Downloads\\Pay Policies - Pay Rules.csv";
		String excelFilePath = "C:\\softwares and jars\\IDD_Automation\\test files\\SOURCE.xlsx";	
		Client_name();
		
		driver.findElement(By.xpath("//em[@class='icon icon-Insight']")).click();
		driver.findElement(By.xpath("//span[text()=' Pay Policies']")).click();
		driver.findElement(By.xpath("//*[text()='Total Number of Pay Rules']")).click();
		driver.findElement(By.xpath("//button[text()=' Export ']")).click();
		System.out.println("HALT FOR 1 SEC");
		Thread.sleep(1000);
		System.out.println("---------------------------------------------");
		XSSFWorkbook workbook = new XSSFWorkbook();
		XSSFSheet sheet = workbook.createSheet();
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
					String trimmedCell = cells[i].replaceAll("^\"|\"$", "");
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
		
		workbook.close();
		System.out.println("Conversion completed: " + excelFilePath);
		String sourceFilePath = "C:\\softwares and jars\\IDD_Automation\\test files\\SOURCE.xlsx";
		String destinationFilePath = "C:\\softwares and jars\\IDD_Automation\\test files\\NAV Tool - IDD Instructions and Client Approval v2.xlsx";
		try (FileInputStream sourceFile = new FileInputStream(sourceFilePath);
				FileInputStream destinationFile = new FileInputStream(destinationFilePath)) {

			// Load the source and destination workbooks
			XSSFWorkbook sourceWorkbook = new XSSFWorkbook(sourceFile);
			XSSFWorkbook destinationWorkbook = new XSSFWorkbook(destinationFile);
			// Get the specific sheet from both workbooks
			XSSFSheet sourceSheet = sourceWorkbook.getSheetAt(0); // Adjust sheet index as needed
			XSSFSheet destinationSheet = destinationWorkbook.getSheetAt(8); // Adjust sheet index as needed
			// Iterate through the rows and cells of the source sheet and copy them to the destination sheet
			int rowCount = sourceSheet.getLastRowNum();

			for (int i = 1; i <= rowCount; i++) {
				Row sourceRow = sourceSheet.getRow(i);
				Row destinationRow = destinationSheet.createRow(i);

				int cellCount = sourceRow.getLastCellNum();
				for (int j = 0; j < cellCount; j++) {
					Cell sourceCell = sourceRow.getCell(j);
					Cell destinationCell = destinationRow.createCell(j);

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
			System.out.println("Pay Rules END");

		} catch (IOException e) {
			e.printStackTrace();
		}
	}

	//------------------------------------------------WORK RULES-----------------------------------------
	public static void Work_Rules() throws InvalidFormatException, InterruptedException, IOException {
		//driver=Base.openBrowser();
		System.out.println("Work_Rule START");
		String csvFilePath = "C:\\Users\\gadhavek\\Downloads\\Pay Policies - Work Rules.csv";
		String excelFilePath = "C:\\softwares and jars\\IDD_Automation\\test files\\SOURCE.xlsx";		
		Client_name();
		
		driver.findElement(By.xpath("//em[@class='icon icon-Insight']")).click();
		driver.findElement(By.xpath("//span[text()=' Pay Policies']")).click();
		driver.findElement(By.xpath("//*[text()='Total Number of Work Rules']")).click();
		driver.findElement(By.xpath("//button[text()=' Export ']")).click();
		System.out.println("HALT FOR 1 SEC");
		Thread.sleep(1000);	
		System.out.println("---------------------------------------------");

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
					String trimmedCell = cells[i].replaceAll("^\"|\"$", "");

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
		
		workbook.close();
		System.out.println("Conversion completed: " + excelFilePath);
		String sourceFilePath = "C:\\softwares and jars\\IDD_Automation\\test files\\SOURCE.xlsx";
		String destinationFilePath = "C:\\softwares and jars\\IDD_Automation\\test files\\NAV Tool - IDD Instructions and Client Approval v2.xlsx";
		try (FileInputStream sourceFile = new FileInputStream(sourceFilePath);
				FileInputStream destinationFile = new FileInputStream(destinationFilePath)) {

			// Load the source and destination workbooks
			XSSFWorkbook sourceWorkbook = new XSSFWorkbook(sourceFile);
			XSSFWorkbook destinationWorkbook = new XSSFWorkbook(destinationFile);

			// Get the specific sheet from both workbooks
			XSSFSheet sourceSheet = sourceWorkbook.getSheetAt(0); // Adjust sheet index as needed
			XSSFSheet destinationSheet = destinationWorkbook.getSheetAt(9); // Adjust sheet index as needed

			// Iterate through the rows and cells of the source sheet and copy them to the destination sheet
			int rowCount = sourceSheet.getLastRowNum();
			//Row row = destinationSheet.createRow(++rowCount); 
			//Cell cell = row.createCell(0);

			for (int i = 1; i <= rowCount; i++) {
				Row sourceRow = sourceSheet.getRow(i);
				Row destinationRow = destinationSheet.createRow(i);

				int cellCount = sourceRow.getLastCellNum();
				for (int j = 0; j < cellCount; j++) {
					Cell sourceCell = sourceRow.getCell(j);
					Cell destinationCell = destinationRow.createCell(j);

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
			System.out.println("Work Rule END");

		} catch (IOException e) {
			e.printStackTrace();
		}
	}

	//-----------------------------------------STANDARD PAY CODES-----------------------------------------
	public static void Standard_Pay_Code() throws InvalidFormatException, InterruptedException, IOException {
		//driver=Base.openBrowser();
		System.out.println("Standard Pay Code START");
		String csvFilePath = "C:\\Users\\gadhavek\\Downloads\\Pay Policies - Pay Codes.csv";
		String excelFilePath = "C:\\softwares and jars\\IDD_Automation\\test files\\SOURCE.xlsx";		
		Client_name();
		
		driver.findElement(By.xpath("//em[@class='icon icon-Insight']")).click();
		driver.findElement(By.xpath("//span[text()=' Pay Policies']")).click();
		driver.findElement(By.xpath("//*[text()='Number of Standard Pay Codes']")).click();
		driver.findElement(By.xpath("//button[text()=' Export ']")).click();
		System.out.println("HALT FOR 1 SEC");
		Thread.sleep(1000);	
		System.out.println("---------------------------------------------");
		//System.out.println(driver.findElement(By.xpath("//div[@class='insights-detail-table']//descendant::div[@class='table-height ng-star-inserted']")).getText());
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
					String trimmedCell = cells[i].replaceAll("^\"|\"$", "");
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
		
		workbook.close();
		System.out.println("Conversion completed: " + excelFilePath);
		String sourceFilePath = "C:\\softwares and jars\\IDD_Automation\\test files\\SOURCE.xlsx";
		String destinationFilePath = "C:\\softwares and jars\\IDD_Automation\\test files\\NAV Tool - IDD Instructions and Client Approval v2.xlsx";
		try (FileInputStream sourceFile = new FileInputStream(sourceFilePath);
				FileInputStream destinationFile = new FileInputStream(destinationFilePath)) {
			// Load the source and destination workbooks
			XSSFWorkbook sourceWorkbook = new XSSFWorkbook(sourceFile);
			XSSFWorkbook destinationWorkbook = new XSSFWorkbook(destinationFile);
			// Get the specific sheet from both workbooks
			XSSFSheet sourceSheet = sourceWorkbook.getSheetAt(0); // Adjust sheet index as needed
			XSSFSheet destinationSheet = destinationWorkbook.getSheetAt(10); // Adjust sheet index as needed
			// Iterate through the rows and cells of the source sheet and copy them to the destination sheet
			int rowCount = sourceSheet.getLastRowNum();
			//Row row = destinationSheet.createRow(++rowCount); 
			//Cell cell = row.createCell(0);
			for (int i = 1; i <= rowCount; i++) {
				Row sourceRow = sourceSheet.getRow(i);
				Row destinationRow = destinationSheet.createRow(i);

				int cellCount = sourceRow.getLastCellNum();
				for (int j = 0; j < cellCount; j++) {
					Cell sourceCell = sourceRow.getCell(j);
					Cell destinationCell = destinationRow.createCell(j);
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
			System.out.println("Standard Pay Codes END");
		} catch (IOException e) {
			e.printStackTrace();
		}
		String deletefile="C:\\Users\\gadhavek\\Downloads\\Pay Policies - Pay Codes.csv";
		File file=new File(deletefile);
		file.delete();
	}
	
	//-----------------------------------------DURATION PAY CODES-----------------------------------------
		public static void Duration_Pay_Code() throws InvalidFormatException, InterruptedException, IOException {
			//driver=Base.openBrowser();
			System.out.println("DURATION PAY CODE START");
			String csvFilePath = "C:\\Users\\gadhavek\\Downloads\\Pay Policies - Pay Codes.csv";
			String excelFilePath = "C:\\softwares and jars\\IDD_Automation\\test files\\SOURCE.xlsx";		
			Client_name();
			
			driver.findElement(By.xpath("//em[@class='icon icon-Insight']")).click();
			driver.findElement(By.xpath("//span[text()=' Pay Policies']")).click();
			driver.findElement(By.xpath("//*[text()='Number of Duration Pay Codes']")).click();
			driver.findElement(By.xpath("//button[text()=' Export ']")).click();
			System.out.println("HALT FOR 1 SEC");
			Thread.sleep(1000);	
			System.out.println("---------------------------------------------");
			//System.out.println(driver.findElement(By.xpath("//div[@class='insights-detail-table']//descendant::div[@class='table-height ng-star-inserted']")).getText());
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
						String trimmedCell = cells[i].replaceAll("^\"|\"$", "");
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
			
			workbook.close();
			System.out.println("Conversion completed: " + excelFilePath);
			String sourceFilePath = "C:\\softwares and jars\\IDD_Automation\\test files\\SOURCE.xlsx";
			String destinationFilePath = "C:\\softwares and jars\\IDD_Automation\\test files\\NAV Tool - IDD Instructions and Client Approval v2.xlsx";
			try (FileInputStream sourceFile = new FileInputStream(sourceFilePath);
					FileInputStream destinationFile = new FileInputStream(destinationFilePath)) {
				// Load the source and destination workbooks
				XSSFWorkbook sourceWorkbook = new XSSFWorkbook(sourceFile);
				XSSFWorkbook destinationWorkbook = new XSSFWorkbook(destinationFile);
				// Get the specific sheet from both workbooks
				XSSFSheet sourceSheet = sourceWorkbook.getSheetAt(0); // Adjust sheet index as needed
				XSSFSheet destinationSheet = destinationWorkbook.getSheetAt(11); // Adjust sheet index as needed
				// Iterate through the rows and cells of the source sheet and copy them to the destination sheet
				int rowCount = sourceSheet.getLastRowNum();
				//Row row = destinationSheet.createRow(++rowCount); 
				//Cell cell = row.createCell(0);
				for (int i = 1; i <= rowCount; i++) {
					Row sourceRow = sourceSheet.getRow(i);
					Row destinationRow = destinationSheet.createRow(i);

					int cellCount = sourceRow.getLastCellNum();
					for (int j = 0; j < cellCount; j++) {
						Cell sourceCell = sourceRow.getCell(j);
						Cell destinationCell = destinationRow.createCell(j);
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
				System.out.println("Duration Pay Codes END");
			} catch (IOException e) {
				e.printStackTrace();
			}
			String deletefile="C:\\Users\\gadhavek\\Downloads\\Pay Policies - Pay Codes.csv";
			File file=new File(deletefile);
			file.delete();
		}
		
		//-----------------------------------------CASCADING PAY CODES-----------------------------------------
		public static void Cascading_Pay_Code() throws InvalidFormatException, InterruptedException, IOException {
			driver=Base.openBrowser();
			System.out.println("Pay_Codes START");
			String csvFilePath = "C:\\Users\\gadhavek\\Downloads\\Pay Policies - Pay Codes.csv";
			String excelFilePath = "C:\\softwares and jars\\IDD_Automation\\test files\\SOURCE.xlsx";
			driver.findElement(By.xpath("//span[contains(text(),'Insights')]")).click();
			driver.findElement(By.xpath("//span[text()=' Pay Policies']")).click();
			driver.findElement(By.xpath("//*[text()='Number of Cascading Pay Codes']")).click();
			
			try {
				driver.findElement(By.xpath("//button[text()=' Export ']")).click();
			} catch (Exception e) {
				System.out.println("No data");
			}
			driver.findElement(By.xpath("//button[text()=' Export ']")).click();
			System.out.println("HALT FOR 1 SEC");
			Thread.sleep(1000);	
			System.out.println("---------------------------------------------");
			//System.out.println(driver.findElement(By.xpath("//div[@class='insights-detail-table']//descendant::div[@class='table-height ng-star-inserted']")).getText());
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
						String trimmedCell = cells[i].replaceAll("^\"|\"$", "");
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
			
			workbook.close();
			System.out.println("Conversion completed: " + excelFilePath);
			String sourceFilePath = "C:\\softwares and jars\\IDD_Automation\\test files\\SOURCE.xlsx";
			String destinationFilePath = "C:\\softwares and jars\\IDD_Automation\\test files\\NAV Tool - IDD Instructions and Client Approval v2.xlsx";
			try (FileInputStream sourceFile = new FileInputStream(sourceFilePath);
					FileInputStream destinationFile = new FileInputStream(destinationFilePath)) {
				// Load the source and destination workbooks
				XSSFWorkbook sourceWorkbook = new XSSFWorkbook(sourceFile);
				XSSFWorkbook destinationWorkbook = new XSSFWorkbook(destinationFile);
				// Get the specific sheet from both workbooks
				XSSFSheet sourceSheet = sourceWorkbook.getSheetAt(0); // Adjust sheet index as needed
				XSSFSheet destinationSheet = destinationWorkbook.getSheetAt(12); // Adjust sheet index as needed
				// Iterate through the rows and cells of the source sheet and copy them to the destination sheet
				int rowCount = sourceSheet.getLastRowNum();
				//Row row = destinationSheet.createRow(++rowCount); 
				//Cell cell = row.createCell(0);
				for (int i = 1; i <= rowCount; i++) {
					Row sourceRow = sourceSheet.getRow(i);
					Row destinationRow = destinationSheet.createRow(i);

					int cellCount = sourceRow.getLastCellNum();
					for (int j = 0; j < cellCount; j++) {
						Cell sourceCell = sourceRow.getCell(j);
						Cell destinationCell = destinationRow.createCell(j);
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
				System.out.println("Pay Codes END");
			} catch (IOException e) {
				e.printStackTrace();
			}
			String deletefile="C:\\Users\\gadhavek\\Downloads\\Pay Policies - Pay Codes.csv";
			File file=new File(deletefile);
			file.delete();
		}
		//-----------------------------------------COMBINED PAY CODES-----------------------------------------
		public static void Combined_Pay_Code() throws InvalidFormatException, InterruptedException, IOException {
			//driver=Base.openBrowser();
			System.out.println("Pay_Codes START");
			String csvFilePath = "C:\\Users\\gadhavek\\Downloads\\Pay Policies - Pay Codes.csv";
			String excelFilePath = "C:\\softwares and jars\\IDD_Automation\\test files\\SOURCE.xlsx";		
			Client_name();
			
			driver.findElement(By.xpath("//em[@class='icon icon-Insight']")).click();
			driver.findElement(By.xpath("//span[text()=' Pay Policies']")).click();
			driver.findElement(By.xpath("//*[text()='Number of Combined Pay Codes']")).click();
			driver.findElement(By.xpath("//button[text()=' Export ']")).click();
			System.out.println("HALT FOR 1 SEC");
			Thread.sleep(1000);	
			System.out.println("---------------------------------------------");
			
			XSSFWorkbook workbook = new XSSFWorkbook();
			XSSFSheet sheet = workbook.createSheet();
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
						String trimmedCell = cells[i].replaceAll("^\"|\"$", "");
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
			
			workbook.close();
			System.out.println("Conversion completed: " + excelFilePath);
			String sourceFilePath = "C:\\softwares and jars\\IDD_Automation\\test files\\SOURCE.xlsx";
			String destinationFilePath = "C:\\softwares and jars\\IDD_Automation\\test files\\NAV Tool - IDD Instructions and Client Approval v2.xlsx";
			try (FileInputStream sourceFile = new FileInputStream(sourceFilePath);
					FileInputStream destinationFile = new FileInputStream(destinationFilePath)) {
				// Load the source and destination workbooks
				XSSFWorkbook sourceWorkbook = new XSSFWorkbook(sourceFile);
				XSSFWorkbook destinationWorkbook = new XSSFWorkbook(destinationFile);
				// Get the specific sheet from both workbooks
				XSSFSheet sourceSheet = sourceWorkbook.getSheetAt(0); // Adjust sheet index as needed
				XSSFSheet destinationSheet = destinationWorkbook.getSheetAt(13); // Adjust sheet index as needed
				// Iterate through the rows and cells of the source sheet and copy them to the destination sheet
				int rowCount = sourceSheet.getLastRowNum();
				//Row row = destinationSheet.createRow(++rowCount); 
				//Cell cell = row.createCell(0);
				for (int i = 1; i <= rowCount; i++) {
					Row sourceRow = sourceSheet.getRow(i);
					Row destinationRow = destinationSheet.createRow(i);

					int cellCount = sourceRow.getLastCellNum();
					for (int j = 0; j < cellCount; j++) {
						Cell sourceCell = sourceRow.getCell(j);
						Cell destinationCell = destinationRow.createCell(j);
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
				System.out.println("Combined Pay Codes END");
			} catch (IOException e) {
				e.printStackTrace();
			}
			String deletefile="C:\\Users\\gadhavek\\Downloads\\Pay Policies - Pay Codes.csv";
			File file=new File(deletefile);
			file.delete();
		}
		//-----------------------------------------EMPLOYMENT TERMS-----------------------------------------
		public static void Employment_Terms() throws InvalidFormatException, InterruptedException, IOException {
			driver=Base.openBrowser();
			System.out.println("Pay_Codes START");
			String csvFilePath = "C:\\Users\\gadhavek\\Downloads\\Pay Policies - Pay Rules.csv";
			String excelFilePath = "C:\\softwares and jars\\IDD_Automation\\test files\\SOURCE.xlsx";		
			Client_name();
		
			driver.findElement(By.xpath("//em[@class='icon icon-Insight']")).click();
			driver.findElement(By.xpath("//span[text()=' Pay Policies']")).click();
			driver.findElement(By.xpath("//*[text()='Total Number of Pay Codes']")).click();
			driver.findElement(By.xpath("//button[text()=' Export ']")).click();
			System.out.println("HALT FOR 1 SEC");
			Thread.sleep(1000);	
			System.out.println("---------------------------------------------");
			//System.out.println(driver.findElement(By.xpath("//div[@class='insights-detail-table']//descendant::div[@class='table-height ng-star-inserted']")).getText());
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
						String trimmedCell = cells[i].replaceAll("^\"|\"$", "");
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
			
			workbook.close();
			System.out.println("Conversion completed: " + excelFilePath);
			String sourceFilePath = "C:\\softwares and jars\\IDD_Automation\\test files\\SOURCE.xlsx";
			String destinationFilePath = "C:\\softwares and jars\\IDD_Automation\\test files\\NAV Tool - IDD Instructions and Client Approval v2.xlsx";
			try (FileInputStream sourceFile = new FileInputStream(sourceFilePath);
					FileInputStream destinationFile = new FileInputStream(destinationFilePath)) {
				// Load the source and destination workbooks
				XSSFWorkbook sourceWorkbook = new XSSFWorkbook(sourceFile);
				XSSFWorkbook destinationWorkbook = new XSSFWorkbook(destinationFile);
				// Get the specific sheet from both workbooks
				XSSFSheet sourceSheet = sourceWorkbook.getSheetAt(0); // Adjust sheet index as needed
				XSSFSheet destinationSheet = destinationWorkbook.getSheetAt(10); // Adjust sheet index as needed
				// Iterate through the rows and cells of the source sheet and copy them to the destination sheet
				int rowCount = sourceSheet.getLastRowNum();
				//Row row = destinationSheet.createRow(++rowCount); 
				//Cell cell = row.createCell(0);
				for (int i = 1; i <= rowCount; i++) {
					Row sourceRow = sourceSheet.getRow(i);
					Row destinationRow = destinationSheet.createRow(i);

					int cellCount = sourceRow.getLastCellNum();
					for (int j = 0; j < cellCount; j++) {
						Cell sourceCell = sourceRow.getCell(j);
						Cell destinationCell = destinationRow.createCell(j);
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
				System.out.println("Employment Pay Codes END");
			} catch (IOException e) {
				e.printStackTrace();
			}
		} 
//-----------------------------------------PAYCODE ACCESS PROFILE--------------------------------------
		
		public static void Paycode_Access_Profile() throws InvalidFormatException, InterruptedException, IOException {
			//driver=Base.openBrowser();
			System.out.println("Pay_Access_Profile START");
			String csvFilePath = "C:\\Users\\gadhavek\\Downloads\\employees_all_columns.csv";
			String excelFilePath = "C:\\softwares and jars\\IDD_Automation\\test files\\SOURCE.xlsx";		
			Client_name();
		
			driver.findElement(By.xpath("//em[@class='icon icon-Migration']")).click();
			driver.findElement(By.xpath("//span[text()=' Employees']")).click();
			Thread.sleep(2000);
			String empcount= driver.findElement(By.xpath("//div[@class='tips d-flex justify-content-start']")).getText();		
			System.out.println(empcount);
			String count[]=empcount.split("are");
			System.out.println(count[1].trim());
			WebElement textElement= driver.findElement(By.xpath("//input[@type='number']"));
			textElement.sendKeys(Keys.BACK_SPACE);
			textElement.sendKeys(Keys.BACK_SPACE);
			textElement.sendKeys(Keys.BACK_SPACE);
			textElement.sendKeys(Keys.BACK_SPACE);
			textElement.sendKeys(count[1]);
			
			driver.findElement(By.xpath("(//*[text()='Download'])[1]")).click();
			System.out.println("HALT FOR 1 SEC");
			Thread.sleep(1000);	
			System.out.println("---------------------------------------------");			
			XSSFWorkbook workbook = new XSSFWorkbook();
			XSSFSheet sheet = workbook.createSheet();
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
						String trimmedCell = cells[i].replaceAll("^\"|\"$", "");
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
			
			workbook.close();
			System.out.println("Conversion completed: " + excelFilePath);
			String sourceFilePath = "C:\\softwares and jars\\IDD_Automation\\test files\\SOURCE.xlsx";
			String destinationFilePath = "C:\\softwares and jars\\IDD_Automation\\test files\\NAV Tool - IDD Instructions and Client Approval v2.xlsx";
			try (FileInputStream sourceFile = new FileInputStream(sourceFilePath);
					FileInputStream destinationFile = new FileInputStream(destinationFilePath)) {
				// Load the source and destination workbooks
				XSSFWorkbook sourceWorkbook = new XSSFWorkbook(sourceFile);
				XSSFWorkbook destinationWorkbook = new XSSFWorkbook(destinationFile);
				// Get the specific sheet from both workbooks
				XSSFSheet sourceSheet = sourceWorkbook.getSheetAt(0); // Adjust sheet index as needed
				XSSFSheet destinationSheet = destinationWorkbook.getSheetAt(15); // Adjust sheet index as needed
				// Iterate through the rows and cells of the source sheet and copy them to the destination sheet
				int [] coltocopy = {64,65};
				//Row row = destinationSheet.createRow(++rowCount); 
				//Cell cell = row.createCell(0);
				for (int i = 0; i <= sourceSheet.getLastRowNum(); i++) {
					Row sourceRow = sourceSheet.getRow(i);
					Row destinationRow = destinationSheet.createRow(i);

					int cellCount = sourceRow.getLastCellNum();
					for (int j = 0; j < cellCount; j++) {
						Cell sourceCell = sourceRow.getCell(j);
						Cell destinationCell = destinationRow.createCell(j);
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
				System.out.println("Employment Pay Codes END");
			} catch (IOException e) {
				e.printStackTrace();
			}
		} 

	public static void main(String[] args) throws InvalidFormatException, InterruptedException, IOException {
		long startTime = System.currentTimeMillis();
		System.out.println("START");
		Fap();
		//Display_profile();
		//Pay_Rules();
		//Work_Rules();
		//Standard_Pay_Code();
		//Duration_Pay_Code();
		//Cascading_Pay_Code();
		//Combined_Pay_Code();
		//Employment_Terms();
		Paycode_Access_Profile();
		long endTime = System.currentTimeMillis();
		System.out.println("-------------------------------------------------------------------------------------------");
		System.out.println();
		System.out.println("It took " + (endTime - startTime)/1000.0 + " SECONDS for the Entire Process.");
		System.out.println("-------------------------------------------------------------------------------------------");

	}
}
