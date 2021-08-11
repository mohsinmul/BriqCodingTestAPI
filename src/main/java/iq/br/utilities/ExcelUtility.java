package iq.br.utilities;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFFormulaEvaluator;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.json.simple.JSONArray;
import org.json.simple.JSONObject;

public class ExcelUtility {

	/*
	 * public static void writeToExcel() throws IOException { String path =
	 * "test.xlsx"; FileInputStream fs = new FileInputStream(path); XSSFWorkbook
	 * workbook = new XSSFWorkbook(fs); Sheet sheet = workbook.getSheetAt(0); int
	 * lastRow = sheet.getLastRowNum(); System.out.println(lastRow); int startRow;
	 * if (lastRow < 0) { // sheet.createRow(startRow) startRow = 0; } else {
	 * startRow = lastRow + 1; }
	 * 
	 * Row row = sheet.createRow(startRow); row.createCell(0).setCellValue("City");
	 * row.createCell(1).setCellValue("Pune");
	 * 
	 * FileOutputStream fos = new FileOutputStream(path); workbook.write(fos);
	 * fos.close(); workbook.close(); }
	 */
	
	public static Map<String, String> readFromExcel(String fileLocation) throws IOException {

		Map data = new HashMap<String, String>();
		FileInputStream fileInputStream = new FileInputStream(fileLocation);
		XSSFWorkbook workbook = new XSSFWorkbook(fileInputStream);
		Sheet sheet = workbook.getSheetAt(0);
		int rowNum = sheet.getLastRowNum();
		for (int i = 0; i <= rowNum; i++) {
			Row row = sheet.getRow(i);
			Cell keyCell = row.getCell(0);
			String key = keyCell.getStringCellValue().trim();
			Cell keyValue = row.getCell(1);
			String value = keyValue.getStringCellValue().trim();
			data.put(key, value);
		}
		return data;
	}
}
