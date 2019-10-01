package com.excel.blank.excelfilereader;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
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
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
public class ReadingBlank {

	public static void main(String[] args) throws IOException {
		FileInputStream file = new FileInputStream(new File("D:/Book4.xlsx"));
		try {
			

			XSSFWorkbook workbook = new XSSFWorkbook(file);

			XSSFSheet sheet = workbook.getSheetAt(0);

			//Iterator<Row> rowIterator = sheet.iterator();

			
			Iterator<Row> rows = sheet.rowIterator();
			while (rows.hasNext()) {
                Row row =  rows.next();
                Iterator<Cell> cells = row.cellIterator();
                while (cells.hasNext()) {
                  Cell cell = cells.next();

                    CellType type = cell.getCellType();
                    if (type == CellType.STRING) {
                        System.out.println("[" + cell.getRowIndex() + ", "
                            + cell.getColumnIndex() + "] = STRING; Value = "
                            + cell.getRichStringCellValue().toString());
                    } else if (type == CellType.NUMERIC) {
                        System.out.println("[" + cell.getRowIndex() + ", "
                            + cell.getColumnIndex() + "] = NUMERIC; Value = "
                            + cell.getNumericCellValue());
                    } else if (type == CellType.BOOLEAN) {
                        System.out.println("[" + cell.getRowIndex() + ", "
                            + cell.getColumnIndex() + "] = BOOLEAN; Value = "
                            + cell.getBooleanCellValue());
                    } else if (type == CellType.BLANK) {
                        System.out.println("[" + cell.getRowIndex() + ", "
                            + cell.getColumnIndex() + "] = BLANK CELL");
                    }
                }
            }
			
		
			file.close();

		} catch (FileNotFoundException e) {
			e.printStackTrace();
		}

		try {

			FileInputStream file1 = new FileInputStream(new File("D:/Book4.xlsx"));

			XSSFWorkbook workbook = new XSSFWorkbook(file1);

			XSSFSheet sheet = workbook.getSheetAt(0);

			
			XSSFCell cell;

			Iterator<Row> rows = sheet.rowIterator();
			ArrayList<Object> cellValues = new ArrayList<Object>();

			while (rows.hasNext()) {
				Row row = rows.next();

				for (int i = 0; i < row.getLastCellNum(); i++) {
					
					//System.out.println( row.getLastCellNum()+" row.getLastCellNum()");

					// cell = row.getCell(i,Row.CREATE_NULL_AS_BLANK );
					cell = (XSSFCell) row.getCell(i);

					String nullresullt = "NaN";

					if (cell != null) {
						
						cellValues.add(cell);
						
						
					} else {
						cellValues.add(nullresullt);

					}

				}

				System.out.println();
			}
			System.out.print(cellValues.toString() + " --> cellValues ");

		} catch (Exception e) {
			e.printStackTrace();
		}
		
		//reedHeaders();
		excelReadForTableCreation();
	
	}
	
	
	
	
	
	
	
	

	public static String reedHeaders() throws FileNotFoundException {
		String headers = null;
		FileInputStream file = new FileInputStream(new File("D:/Book3.xlsx"));
		try {
			DataFormatter dataFormatter = new DataFormatter();
			List<String> cellValues = new ArrayList<String>();
			XSSFWorkbook workbook = new XSSFWorkbook(file);

			 Iterator<Sheet> sheetIterator = workbook.sheetIterator();
			 
			while (sheetIterator.hasNext()) {
				Sheet sheet = sheetIterator.next();
				Iterator<Row> rowIterator = sheet.rowIterator();
				while (rowIterator.hasNext()) {
					Row row = rowIterator.next();

					// Now let's iterate over the columns of the current row
					Iterator<Cell> cellIterator = row.cellIterator();

					while (cellIterator.hasNext()) {
						Cell cell = cellIterator.next();
						String cellValue = dataFormatter.formatCellValue(cell);
						cellValues.add(cellValue);
					}
					break;
				}
				//System.out.println(cellValues);
			}
			headers = String.join(",", cellValues);
			System.out.println("headers ---->"+headers);
			

		} catch (Exception e) {
			e.printStackTrace();
		}
		if (!headers.isEmpty())
			return headers;
		else
			return null;
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

					while (cellIterator.hasNext()) {

						Cell cell = cellIterator.next();

			
						CellType type = cell.getCellType();

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
							// rows.add(Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
							// rows.add(null);
						} else {
							rows.add(dataFormatter.formatCellValue(cell));

						}

					}
					sheetList.add(rows);

				}System.out.println("sheetList"+sheetList);

				sheets.add(sheetList);

			}

			
System.out.println("sheets--->"+sheets);
			return sheets;
		} catch (Exception e) {
			e.printStackTrace();
		}
		return null;
	}
	
	
	
	
	
	
}
