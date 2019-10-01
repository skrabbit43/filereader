/*
 * package com.excel.blank.excelfilereader;
 * 
 * import java.io.File; import java.io.FileInputStream; import
 * java.io.FileNotFoundException; import java.io.IOException; import
 * java.util.Iterator;
 * 
 * import org.apache.poi.ss.usermodel.Cell; import
 * org.apache.poi.ss.usermodel.Row; import
 * org.apache.poi.xssf.usermodel.XSSFSheet; import
 * org.apache.poi.xssf.usermodel.XSSFWorkbook;
 * 
 * public class ReadingFile {
 * 
 * public static void main(String[] args) throws IOException {
 * 
 * try { FileInputStream file = new FileInputStream(new File("D:/sample.xlsx"));
 * 
 * XSSFWorkbook workbook = new XSSFWorkbook(file);
 * 
 * XSSFSheet sheet = workbook.getSheetAt(0);
 * 
 * Iterator<Row> rowIterator = sheet.iterator();
 * 
 * while (rowIterator.hasNext()) { Row row = rowIterator.next(); Iterator<Cell>
 * cellIterator = row.cellIterator(); while (cellIterator.hasNext()) { Cell cell
 * = cellIterator.next(); switch (cell.getCellType()) { case
 * Cell.CELL_TYPE_NUMERIC: System.out.print(cell.getNumericCellValue() + " ");
 * break; case Cell.CELL_TYPE_STRING: System.out.print(cell.getStringCellValue()
 * + " "); break; } } System.out.println("\n"); } file.close();
 * 
 * } catch (FileNotFoundException e) { e.printStackTrace(); }
 * 
 * }
 * 
 * }
 */