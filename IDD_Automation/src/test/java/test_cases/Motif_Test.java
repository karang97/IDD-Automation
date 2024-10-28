package test_cases;
import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.IOException;
import java.io.InputStreamReader;
import java.util.Arrays;
import java.util.Comparator;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;
import base.Base;

public class Motif_Test extends Base{

	public static void queryaccess() {
		driver.findElement(By.xpath("//span[text()='Client Tools']")).click();
		driver.findElement(By.xpath("//div[@id='app_singleClientQueryLaunch']")).click();
		driver.findElement(By.xpath("//textarea[@id='crossClientQuery_codeContainer']")).click();
	}
	
	public static void execute_query() throws InterruptedException {
		driver.findElement(By.xpath("//textarea[@autocorrect='off']")).sendKeys(Keys.ENTER);
		WebElement ok=driver.findElement(By.xpath("//span[@class='ui-button-text' and text()='OK']"));
		ok.click();
		WebElement execute=driver.findElement(By.xpath("//span[@class='ui-button-text'and text()='Execute']"));
		execute.click();
		Thread.sleep(2000);
		driver.findElement(By.xpath("//*[text()='Results']")).click();
	}
	
	public static void download_file() throws InterruptedException {
		//Complete
		WebDriverWait wait = new WebDriverWait(driver,30);
				WebElement completeElement=driver.findElement(By.xpath("//div[@id='crossClientQuery_mainDialog_results_tab']//tbody//tr[1]//td[@class='status requestList_statusColumn']"));
				WebElement completetext=driver.findElement(By.xpath("//div[@id='crossClientQuery_mainDialog_results_tab']//tbody//tr[1]//td[.='Complete']"));				
				wait.until(ExpectedConditions.visibilityOf(completetext));
				driver.findElement(By.xpath("//div[@id='crossClientQuery_mainDialog_results_tab']//tbody//tr[1]//td[@class='actions requestList_actionsColumn']//a")).click();	
	}

	//-------------------------------------Wage_Profile()--------------------------------------

	public static void Wage_Profile() throws InvalidFormatException, InterruptedException, IOException {
		driver=Base.openMotif();
		XSSFWorkbook workbook = new XSSFWorkbook(new File("C:\\softwares and jars\\IDD_Automation\\test files\\motif.xlsx"));
		XSSFSheet motifSheet  = workbook.getSheetAt(0);
		queryaccess();
		for (int i = 0; i <=motifSheet.getLastRowNum();i++) {			
			driver.findElement(By.xpath("//textarea[@autocorrect='off']")).sendKeys(motifSheet.getRow(i).getCell(0).toString().trim());						
		}		
		execute_query();
		Thread.sleep(2000);
		download_file();
		Thread.sleep(2000);
		//Rename the File	
		String downloadDirPath= "C:/Users/gadhavek/Downloads";
		File downloadDir= new File(downloadDirPath);
		File[] files = downloadDir.listFiles((dir,name)-> name.startsWith("7"));
		Arrays.sort(files, Comparator.comparingLong(File:: lastModified).reversed());
		File latestFile= files[0];
		String newFileName="Wage_Profile" + ".csv";
		String newFilePath= downloadDirPath + "/" + newFileName;
		File newFile=new File(newFilePath);
		latestFile.renameTo(newFile);
		Thread.sleep(1000);	
		String csvFilePath = "C:\\Users\\gadhavek\\Downloads\\Wage_Profile.csv";
		String excelFilePath = "C:\\softwares and jars\\IDD_Automation\\test files\\SOURCE.xlsx";
		//csv to Excel
		
		XSSFWorkbook workbook1 = new XSSFWorkbook();
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
		// Close the workbook
		workbook.close();
		
		String sourceFilePath = "C:\\softwares and jars\\IDD_Automation\\test files\\motif.xlsx";
		String destinationFilePath = "C:\\softwares and jars\\IDD_Automation\\test files\\NAV Tool - IDD Instructions and Client Approval v2.xlsx";
		try (FileInputStream sourceFile = new FileInputStream(sourceFilePath);
				FileInputStream destinationFile = new FileInputStream(destinationFilePath)) {
			// Load the source and destination workbooks
			XSSFWorkbook sourceWorkbook = new XSSFWorkbook(sourceFile);
			XSSFWorkbook destinationWorkbook = new XSSFWorkbook(destinationFile);
			// Get the specific sheet from both workbooks
			XSSFSheet sourceSheet = sourceWorkbook.getSheetAt(4); // Adjust sheet index as needed
			XSSFSheet destinationSheet = destinationWorkbook.getSheetAt(20); // Adjust sheet index as needed
			// Iterate through the rows and cells of the source sheet and copy them to the destination sheet
			int rowCount = sourceSheet.getLastRowNum();
			//Row row = destinationSheet.createRow(++rowCount); 
			//Cell cell = row.createCell(0);
			for (int i = 1; i <= rowCount; i++) {
				Row sourceRow = sourceSheet.getRow(i);
				Row destinationRow = destinationSheet.createRow(i);

				int cellCount = sourceRow.getLastCellNum();
				for (int j = 4; j < cellCount; j++) {
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
			System.out.println("Wage Profile Data copied successfully!");
		} catch (IOException e) {
			e.printStackTrace();
		}						
	}
	//-------------------------------------Clock_Query()--------------------------------------
	public static void Clock_Query() throws InvalidFormatException, InterruptedException, IOException {
		Thread.sleep(2000);
		driver.findElement(By.xpath("//span[@class='ui-button-icon-primary ui-icon ui-icon-closethick']")).click();
		Thread.sleep(2000);
		XSSFWorkbook workbook = new XSSFWorkbook(new File("C:\\softwares and jars\\IDD_Automation\\test files\\motif.xlsx"));
		XSSFSheet motifSheet  = workbook.getSheetAt(1);
		queryaccess();
		for (int i = 0; i <=motifSheet.getLastRowNum();i++) {			
			driver.findElement(By.xpath("//textarea[@autocorrect='off']")).sendKeys(motifSheet.getRow(i).getCell(0).toString().trim());						
		}
		execute_query();
		Thread.sleep(2000);
		download_file();
		Thread.sleep(2000);
		//Rename the File
		String downloadDirPath= "C:/Users/gadhavek/Downloads";
		File downloadDir= new File(downloadDirPath);
		File[] files = downloadDir.listFiles((dir,name)-> name.startsWith("7"));
		Arrays.sort(files, Comparator.comparingLong(File:: lastModified).reversed());
		File latestFile= files[0];
		String newFileName="Clock_Query" + ".csv";
		String newFilePath= downloadDirPath + "/" + newFileName;
		File newFile=new File(newFilePath);
		latestFile.renameTo(newFile);
		Thread.sleep(1000);	
		String csvFilePath = "C:\\Users\\gadhavek\\Downloads\\Clock_Query.csv";
		String excelFilePath = "C:\\softwares and jars\\IDD_Automation\\test files\\SOURCE.xlsx";
		//csv to Excel
		
		XSSFWorkbook workbook1 = new XSSFWorkbook();
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
		// Close the workbook
		workbook.close();
		
		String sourceFilePath = "C:\\softwares and jars\\IDD_Automation\\test files\\motif.xlsx";
		String destinationFilePath = "C:\\softwares and jars\\IDD_Automation\\test files\\NAV Tool - IDD Instructions and Client Approval v2.xlsx";
		try (FileInputStream sourceFile = new FileInputStream(sourceFilePath);
				FileInputStream destinationFile = new FileInputStream(destinationFilePath)) {
			// Load the source and destination workbooks
			XSSFWorkbook sourceWorkbook = new XSSFWorkbook(sourceFile);
			XSSFWorkbook destinationWorkbook = new XSSFWorkbook(destinationFile);
			// Get the specific sheet from both workbooks
			XSSFSheet sourceSheet = sourceWorkbook.getSheetAt(4); // Adjust sheet index as needed
			XSSFSheet destinationSheet = destinationWorkbook.getSheetAt(22); // Adjust sheet index as needed
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
			System.out.println("Clock Query Data copied successfully!");
		} catch (IOException e) {
			e.printStackTrace();
		}
	}
	//-------------------------------------Accural_Profile()--------------------------------------

	public static void Accural_Profile() throws InvalidFormatException, InterruptedException, IOException {
		Thread.sleep(2000);
		driver.findElement(By.xpath("//span[@class='ui-button-icon-primary ui-icon ui-icon-closethick']")).click();
		Thread.sleep(2000);
		XSSFWorkbook workbook = new XSSFWorkbook(new File("C:\\files\\motif.xlsx"));
		XSSFSheet motifSheet  = workbook.getSheetAt(2);
		queryaccess();
		for (int i = 0; i <=motifSheet.getLastRowNum();i++) {			
			driver.findElement(By.xpath("//textarea[@autocorrect='off']")).sendKeys(motifSheet.getRow(i).getCell(0).toString().trim());						
		}
		
		execute_query();
		Thread.sleep(2000);
		download_file();
		
Thread.sleep(2000);
		
		//Rename the File
String downloadDirPath= "C:/Users/gadhavek/Downloads";
File downloadDir= new File(downloadDirPath);
File[] files = downloadDir.listFiles((dir,name)-> name.startsWith("7"));
Arrays.sort(files, Comparator.comparingLong(File:: lastModified).reversed());
File latestFile= files[0];
String newFileName="Accural_Profile" + ".csv";
String newFilePath= downloadDirPath + "/" + newFileName;
File newFile=new File(newFilePath);
latestFile.renameTo(newFile);
Thread.sleep(1000);		
		String csvFilePath = "C:\\Users\\gadhavek\\Downloads\\Accural_Profile.csv";
		String excelFilePath = "C:\\softwares and jars\\IDD_Automation\\test files\\SOURCE.xlsx";
		//csv to Excel
		
		XSSFWorkbook workbook1 = new XSSFWorkbook();
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
		// Close the workbook
		workbook.close();
		
		String sourceFilePath = "C:\\softwares and jars\\IDD_Automation\\test files\\motif.xlsx";
		String destinationFilePath = "C:\\softwares and jars\\IDD_Automation\\test files\\NAV Tool - IDD Instructions and Client Approval v2.xlsx";
		try (FileInputStream sourceFile = new FileInputStream(sourceFilePath);
				FileInputStream destinationFile = new FileInputStream(destinationFilePath)) {
			// Load the source and destination workbooks
			XSSFWorkbook sourceWorkbook = new XSSFWorkbook(sourceFile);
			XSSFWorkbook destinationWorkbook = new XSSFWorkbook(destinationFile);
			// Get the specific sheet from both workbooks
			XSSFSheet sourceSheet = sourceWorkbook.getSheetAt(4); // Adjust sheet index as needed
			XSSFSheet destinationSheet = destinationWorkbook.getSheetAt(24); // Adjust sheet index as needed
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
			System.out.println("Accural profile Data copied successfully!");
		} catch (IOException e) {
			e.printStackTrace();
		}
		
	}
	//-------------------------------------Activity_Profile()--------------------------------------

	public static void Activity_Profile() throws InvalidFormatException, InterruptedException, IOException {
		Thread.sleep(2000);
		driver.findElement(By.xpath("//span[@class='ui-button-icon-primary ui-icon ui-icon-closethick']")).click();
		Thread.sleep(2000);
		XSSFWorkbook workbook = new XSSFWorkbook(new File("C:\\softwares and jars\\IDD_Automation\\test files\\motif.xlsx"));
		XSSFSheet motifSheet  = workbook.getSheetAt(3);
		queryaccess();
		for (int i = 0; i <=motifSheet.getLastRowNum();i++) {			
			driver.findElement(By.xpath("//textarea[@autocorrect='off']")).sendKeys(motifSheet.getRow(i).getCell(0).toString().trim());						
		}
		execute_query();
		Thread.sleep(2000);
		download_file();
		
Thread.sleep(2000);
		
		//Rename the File
String downloadDirPath= "C:/Users/gadhavek/Downloads";
File downloadDir= new File(downloadDirPath);
File[] files = downloadDir.listFiles((dir,name)-> name.startsWith("7"));
Arrays.sort(files, Comparator.comparingLong(File:: lastModified).reversed());
File latestFile= files[0];
String newFileName="Activity_Profile" + ".csv";
String newFilePath= downloadDirPath + "/" + newFileName;
File newFile=new File(newFilePath);
latestFile.renameTo(newFile);
Thread.sleep(1000);	
		String csvFilePath = "C:\\Users\\gadhavek\\Downloads\\Activity_Profile.csv";
		String excelFilePath = "C:\\softwares and jars\\IDD_Automation\\test files\\SOURCE.xlsx";
		//csv to Excel
		
		XSSFWorkbook workbook1 = new XSSFWorkbook();
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
		
		String sourceFilePath = "C:\\softwares and jars\\IDD_Automation\\test files\\motif.xlsx";
		String destinationFilePath = "C:\\softwares and jars\\IDD_Automation\\test files\\NAV Tool - IDD Instructions and Client Approval v2.xlsx";
		try (FileInputStream sourceFile = new FileInputStream(sourceFilePath);
				FileInputStream destinationFile = new FileInputStream(destinationFilePath)) {
			// Load the source and destination workbooks
			XSSFWorkbook sourceWorkbook = new XSSFWorkbook(sourceFile);
			XSSFWorkbook destinationWorkbook = new XSSFWorkbook(destinationFile);
			// Get the specific sheet from both workbooks
			XSSFSheet sourceSheet = sourceWorkbook.getSheetAt(4); // Adjust sheet index as needed
			XSSFSheet destinationSheet = destinationWorkbook.getSheetAt(25); // Adjust sheet index as needed
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
			System.out.println("Activity Profile Data copied successfully!");
		} catch (IOException e) {
			e.printStackTrace();
		}
	}
	public static void AdjRule_PS() throws InvalidFormatException, InterruptedException, IOException {
//		XSSFWorkbook Credentials = new XSSFWorkbook(new File("C:\\softwares and jars\\IDD_Automation\\test files\\login_details.xlsx"));
//		XSSFSheet Sheet = Credentials.getSheetAt(0);
//		Thread.sleep(2000);
//		driver.findElement(By.xpath("//span[@class='ui-button-icon-primary ui-icon ui-icon-closethick']")).click();
//		Thread.sleep(1000);
//		driver.findElement(By.xpath("//span[normalize-space()='Request Account']")).click();
//		String eTimeNameString = driver.findElement(By.xpath("")).getText();
//		String dataCentre = driver.findElement(By.xpath("")).getText();
//		String accountName= Sheet.getRow(1).getCell(2).toString().trim();
//		String password = driver.findElement(By.xpath("//div[@role='dialog']//tr[@rowid='password']//td[@class='Value']")).getText();
//		
		
		try {
            // PowerShell command
            ProcessBuilder builder = new ProcessBuilder("C:\\softwares and jars\\IDD_Automation\\test files\\AdjRule.ps1", "-Command", "Get-Process");
            // Redirect the error stream to capture errors
            builder.redirectErrorStream(true);
           // Start the process
            Process process = builder.start();
           // Read the output
            BufferedReader reader = new BufferedReader(new InputStreamReader(process.getInputStream()));
            String line;
            while ((line = reader.readLine()) != null) {
                System.out.println(line);  // Output the result
            }
            // Wait for the process to finish and get the exit code
            int exitCode = process.waitFor();
            System.out.println("Exit Code: " + exitCode);

        } catch (Exception e) {
            e.printStackTrace();
        }

	}
	
	public static void main(String[] args) throws InvalidFormatException, InterruptedException, IOException {
		long startTime = System.currentTimeMillis();
		Wage_Profile();
		Clock_Query();
		Accural_Profile();
		Activity_Profile();
		long endTime = System.currentTimeMillis();
		System.out.println("-------------------------------------------------------------------------------------------");
		System.out.println("It took " + (endTime - startTime)/1000.0 + " SECONDS for the Entire Process.");
		System.out.println("-------------------------------------------------------------------------------------------");


	}	
}

