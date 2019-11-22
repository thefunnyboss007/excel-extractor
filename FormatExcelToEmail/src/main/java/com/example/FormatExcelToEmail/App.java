package com.example.FormatExcelToEmail;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class App {
	private static final String FILE_NAME = "dummy.xlsx";

	// private static final String FILE_NAME = "/tmp/MyFirstExcel.xlsx";

	public static void main(String[] args) throws IOException {

		FileInputStream file = new FileInputStream(new File(FILE_NAME));
		XSSFWorkbook workbook = new XSSFWorkbook(file);
		XSSFSheet sheet = workbook.getSheetAt(0);

		int rowCount = sheet.getLastRowNum();
		System.out.println(rowCount);

		Cell cell = null;

		for (int i = 0; i < rowCount; i++) {
			cell = sheet.getRow(i).getCell(1);
			cell.setCellValue("Cat");
			System.out.println(cell);
		}

		file.close();

		FileOutputStream outputFile = new FileOutputStream(new File(FILE_NAME));
		workbook.write(outputFile);
		outputFile.close();

	}
}
