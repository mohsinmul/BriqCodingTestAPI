package iq.br.development;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;
import java.util.Map;
import java.util.concurrent.TimeUnit;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.json.simple.JSONObject;
import org.testng.annotations.Test;

import io.restassured.RestAssured;
import io.restassured.http.Header;
import io.restassured.response.ExtractableResponse;
import io.restassured.response.Response;
import iq.br.utilities.ExcelUtility;

public class SaveAndRetrieveAPIResponse {
	
	@Test
	public void captureAndSaveDataInExcel() {
		ExtractableResponse<Response> response = RestAssured.given()
				.get("https://data.sfgov.org/resource/p4e4-a5a7.json").then().extract();
		System.out.println(response.headers().size());
		String filePath = "Leads.xlsx";
		FileInputStream fs = null;
		try {
			fs = new FileInputStream(filePath);
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		}
		try {
			XSSFWorkbook workbook = new XSSFWorkbook(fs);
			Sheet sheet = workbook.getSheetAt(0);
			int lastRow = sheet.getLastRowNum();
			System.out.println(lastRow);
			int startRow;
			if (lastRow < 0)
				startRow = 0;
			else
				startRow = lastRow + 1;

			Row row;

			List<Header> headerList = response.headers().asList();
			for (Header header : headerList) {
				row = sheet.createRow(startRow);
				row.createCell(0).setCellValue(header.getName());
				row.createCell(1).setCellValue(header.getValue());
				startRow++;
			}
			//response body not able to save in excel cause of cell limit.
			// System.out.println("ResponseBody : ");
			// System.out.println(response.body().asPrettyString());

			row = sheet.createRow(startRow);
			row.createCell(0).setCellValue("Status Code");
			row.createCell(1).setCellValue(String.valueOf(response.statusCode()));
			startRow++;

			row = sheet.createRow(startRow);
			row.createCell(0).setCellValue("Status Line");
			row.createCell(1).setCellValue(response.statusLine());
			startRow++;

			row = sheet.createRow(startRow);
			row.createCell(0).setCellValue("Content Type");
			row.createCell(1).setCellValue(response.contentType());
			startRow++;

			row = sheet.createRow(startRow);
			row.createCell(0).setCellValue("Time");
			row.createCell(1).setCellValue(String.valueOf(response.timeIn(TimeUnit.SECONDS)));
			startRow++;

			FileOutputStream fos = new FileOutputStream(filePath);
			workbook.write(fos);
			fos.close();
			workbook.close();
			System.out.println("Successfully saved data...");
		} catch (IOException e) {
			e.printStackTrace();
		}
	}
	
	@Test
	public void readDataFromExcel() {
		try {
			Map<String, String> readData = ExcelUtility.readFromExcel("Leads.xlsx");
			JSONObject obj = new JSONObject(readData);
			System.out.println("Data retrieved from Leads.xlsx : ");
			System.out.println(obj.toJSONString());
		} catch (IOException e) {
			e.printStackTrace();
		}
	}
}
