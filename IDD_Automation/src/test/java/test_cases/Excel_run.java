package test_cases;

import java.io.BufferedReader;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.IOException;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Excel_run {
	public static void convertCSVToExcel(String csvFilePath, String excelFilePath) throws IOException {
		// Create a new Workbook
		XSSFWorkbook workbook = new XSSFWorkbook();
		// Create a new sheet
		XSSFSheet sheet = workbook.createSheet("CSV_Data");

		// Read the CSV file
		try (BufferedReader br = new BufferedReader(new FileReader(csvFilePath))) {
			String line;
			int rowNum = 0;
			// Loop through each line of the CSV file
			while ((line = br.readLine()) != null) {
				// Split the line by commas to get individual values
				String[] values = line.split(",");
				// Create a new row in the Excel sheet
				Row row = sheet.createRow(rowNum++);

				// Loop through the values and add them to the row
				for (int i = 0; i < values.length; i++) {
					row.createCell(i).setCellValue(values[i]);
				}
			}
		}

		// Write the workbook to the output Excel file
		try (FileOutputStream fos = new FileOutputStream(excelFilePath)) {
			workbook.write(fos);
		}

		// Close the workbook
		workbook.close();
		System.out.println("CSV data has been written to " + excelFilePath);
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

		} catch (IOException e) {
			e.printStackTrace();
		}
	}





	public static void main(String[] args) throws IOException, InvalidFormatException {
		String csvFilePath = "C:\\Users\\gadhavek\\Downloads\\Persona Components - Function Access Profiles.csv";
		String excelFilePath = "C:\\files\\SOURCE.xlsx";

		convertCSVToExcel(csvFilePath, excelFilePath);
		ExcelCopy();

	}
}


