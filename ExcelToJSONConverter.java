package com.javacodepoint.example;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.node.ArrayNode;
import com.fasterxml.jackson.databind.node.ObjectNode;


public class ExcelToJSONConverter {

	private ObjectMapper mapper = new ObjectMapper();


	 //Method to convert excel sheet data to JSON format

	public JsonNode excelToJson(File excel) {
		// hold the excel data sheet wise
		ObjectNode excelData = mapper.createObjectNode();
		FileInputStream fis = null;
		Workbook workbook = null;
		try {
			// Creating file input stream
			fis = new FileInputStream(excel);

			String filename = excel.getName().toLowerCase();
			if (filename.endsWith(".xls") || filename.endsWith(".xlsx")) {
				// creating workbook object based on excel file format
				if (filename.endsWith(".xls")) {
					workbook = new HSSFWorkbook(fis);
				} else {
					workbook = new XSSFWorkbook(fis);
				}

				// Reading each sheet one by one
				for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
					Sheet sheet = workbook.getSheetAt(i);
					String sheetName = sheet.getSheetName();

					List<String> headers = new ArrayList<String>();
					ArrayNode sheetData = mapper.createArrayNode();
					// Reading each row of the sheet
					for (int j = 0; j <= sheet.getLastRowNum(); j++) {
						Row row = sheet.getRow(j);
						if (j == 0) {
							// reading sheet header's name
							for (int k = 0; k < row.getLastCellNum(); k++) {
								headers.add(row.getCell(k).getStringCellValue());
							}
						} else {
							// reading work sheet data
							ObjectNode rowData = mapper.createObjectNode();
							for (int k = 0; k < headers.size(); k++) {
								Cell cell = row.getCell(k);
								String headerName = headers.get(k);
								if (cell != null) {
									switch (cell.getCellType()) {
									case FORMULA:
										rowData.put(headerName, cell.getCellFormula());
										break;
									case BOOLEAN:
										rowData.put(headerName, cell.getBooleanCellValue());
										break;
									case NUMERIC:
										rowData.put(headerName, cell.getNumericCellValue());
										break;
									case BLANK:
										rowData.put(headerName, "");
										break;
									default:
										rowData.put(headerName, cell.getStringCellValue());
										break;
									}
								} else {
									rowData.put(headerName, "");
								}
							}
							sheetData.add(rowData);
						}
					}
					excelData.set(sheetName, sheetData);
				}
				return excelData;
			} else {
				throw new IllegalArgumentException("File format not supported.");
			}
		} catch (Exception e) {
			e.printStackTrace();
		} finally {
			if (workbook != null) {
				try {
					workbook.close();
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
			if (fis != null) {
				try {
					fis.close();
				} catch (IOException e) {
					e.printStackTrace();
				}
			}

		}
		return null;
	}


	 //Method to write the json data into file

	public boolean writeJsonToFile(File jsonFile, JsonNode jsonData) {
		try {
			if (jsonFile.getName().endsWith(".json")) {
				if (!jsonFile.exists()) {
					jsonFile.createNewFile();
				}
				FileWriter fw = new FileWriter(jsonFile);
				fw.write(jsonData.toPrettyString());
				fw.close();
				return true;
			} else {
				throw new IllegalArgumentException("File should be .json file only");
			}
		} catch (Exception e) {
			e.printStackTrace();
		}
		return false;
	}


	 //Main method to test

	public static void main(String[] args) {

		// Creating a file object with specific file path
		File excel = new File("/Users/mac/Desktop/Professional Life/Global Logic/Child product fields.xlsx");
		ExcelToJSONConverter converter = new ExcelToJSONConverter();
		JsonNode data = converter.excelToJson(excel);
		System.out.println("Excel file contains the Data:\n" + data);

		File jsonFile = new File("/Users/mac/Desktop/Professional Life/Global Logic/ApachePOIExcelExample/src/main/java/package.json");
		// write the json object into file
		if (converter.writeJsonToFile(jsonFile, data)) {
			System.out.println("JSON file has been created.");
		}
	}
}
