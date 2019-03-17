package com.malhotra.file;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class FileRead {

	private static final String REMOVE = "F:/remove.xlsx";
	private static final String ART = "F:/art.xlsx";
	private static final String FINAL_EXCEL = "F:/final_excel.xlsx";
	
	public static void main(String[] args) {
		List<String> data = new ArrayList<>();
		data.add("a00");
		System.out.println(data);
		try {
			Workbook workbook = WorkbookFactory.create(new File(REMOVE));
		} catch (EncryptedDocumentException | InvalidFormatException | IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

	}

	/*public static void main(String[] args) throws InvalidFormatException, IOException {
		
		System.out.println(REMOVE);
		
		List<String> data = new ArrayList<>();
		data.add("a00");
		System.out.println(data);
		try {
			Workbook workbook = WorkbookFactory.create(new File(REMOVE));
			System.out.print(workbook);
			Sheet sheet = workbook.getSheetAt(0);
			for(Row row : sheet) {
				for(Cell cell : row) {
					String name = cell.getStringCellValue().trim();
					data.add(name);
					//String email = getEmailIdFromExcel(name);
					data.add(email);
					if(!email.equals(""))
						writeToNewSheet(data);
				}
				System.out.println(data);
				data.clear();
			}
			//workbook.close();
		} catch (EncryptedDocumentException e) {
			e.printStackTrace();
		}

	}

	private static String getEmailIdFromExcel(String name) throws EncryptedDocumentException, InvalidFormatException, IOException {
		String email = "";
		Workbook workbook = WorkbookFactory.create(new File(ART));
		Sheet sheet = workbook.getSheetAt(0);
		for(Row row : sheet) {
			Cell cell = row.getCell(0);
			String cellValue = cell.getStringCellValue().trim();
			if(name.equalsIgnoreCase(cellValue)) {
				email = row.getCell(1).toString().trim();
			}
		}
		workbook.close();
		return email;
	}

	private static void writeToNewSheet(List<String> data) throws EncryptedDocumentException, InvalidFormatException, IOException {

		FileInputStream fileInputStream = new FileInputStream(FINAL_EXCEL);
		Workbook workbook = WorkbookFactory.create(fileInputStream);
		Sheet sheet = workbook.getSheetAt(0);
		int rowNum = sheet.getLastRowNum();
		Row row = sheet.createRow(++rowNum);
		int colNum = 0;
		for(String string : data) {
			Cell cell = row.createCell(colNum++);
			cell.setCellValue(string);
		}
		sheet.autoSizeColumn(0);
		sheet.autoSizeColumn(1);
		fileInputStream.close();
		FileOutputStream fileOutputStream = new FileOutputStream(FINAL_EXCEL);
		workbook.write(fileOutputStream);
		workbook.close();
		fileOutputStream.close();	
	}*/

}
