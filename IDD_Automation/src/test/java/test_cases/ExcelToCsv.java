package test_cases;

import java.io.BufferedReader;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.IOException;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelToCsv {
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
	public static void main(String[] args) throws IOException, InvalidFormatException {
		String csvFilePath = "C:\\files\\Test_data_CSV.csv";
		String excelFilePath = "C:\\files\\Test_target.xlsx";

		convertCSVToExcel(csvFilePath, excelFilePath);
	}

}
