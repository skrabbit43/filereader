package com.excel.blank.excelfilereader;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ObtainingCellType {
	public static void main(String[] args) throws IOException {
		excelReadForTableCreation();
	}
	
	

	public static List<List<List<Object>>> excelReadForTableCreation() throws IOException {
		FileInputStream file = new FileInputStream(new File("D:/Book4.xlsx"));
		try {
			DataFormatter dataFormatter = new DataFormatter();
			XSSFWorkbook workbook = new XSSFWorkbook(file);

			Iterator<Sheet> sheetIterator = workbook.sheetIterator();
			List<List<List<Object>>> sheets = new ArrayList<List<List<Object>>>();

			while (sheetIterator.hasNext()) {
				List<List<Object>> sheetList = new ArrayList<List<Object>>();
				Sheet sheet = sheetIterator.next();
				Iterator<Row> rowIterator = sheet.rowIterator();
				while (rowIterator.hasNext()) {
					Row row = rowIterator.next();
					List<Object> rows = new ArrayList<Object>();
					
					
					Iterator<Cell> cellIterator = row.cellIterator();
					ArrayList<Object> cellValues = new ArrayList<Object>();

					while (cellIterator.hasNext()) {

						Cell cell = cellIterator.next();

						CellType type = cell.getCellType();
						
						
						
						
						if (type == CellType._NONE) {
							rows.add(null);

						} 
						
						
						if (type == CellType.STRING) {
							rows.add(cell.getRichStringCellValue().toString());

						} else if (type == CellType.NUMERIC) {

							if (HSSFDateUtil.isCellDateFormatted(cell)) {
								rows.add(cell.getDateCellValue());
							} else if (dataFormatter.formatCellValue(cell).contains(".")) {
								try {
									rows.add(Double.parseDouble(dataFormatter.formatCellValue(cell)));
								} catch (Exception e) {
									rows.add(cell.getRichStringCellValue().toString());
								}
							} else {
								try {
									rows.add(Long.parseLong(dataFormatter.formatCellValue(cell)));
								} catch (Exception e) {
									rows.add(dataFormatter.formatCellValue(cell));
								}
							}

						} else if (type == CellType.BOOLEAN) {
							rows.add(cell.getBooleanCellValue());
						} else if (type == CellType.BLANK) {
							rows.add(dataFormatter.formatCellValue(cell));
						} else {
							rows.add(dataFormatter.formatCellValue(cell));

						}

					}
					sheetList.add(rows);
					for (int i = 0; i < row.getLastCellNum(); i++) {
						Cell c = row.getCell(i);
						if (c == null) {
							rows.add(Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
							sheetList.add(rows);break;
						}

					}

				}
				System.out.println("sheetList " + "\n" + sheetList);

				sheets.add(sheetList);

			}

			System.out.println("sheets--->" + sheets);
			return sheets;
		} catch (Exception e) {
			e.printStackTrace();
		}
		return null;
	}
	
	/*
	 * public static List<List<Object>> adjustColumnCount(List<List<Object>>
	 * shetData, int listSize) {
	 * 
	 * for (int i = 0; i < shetData.size(); i++) { for (int j =
	 * shetData.get(i).size(); j < listSize; j++) { shetData.get(j).add(null); } }
	 * return shetData; }
	 */


}