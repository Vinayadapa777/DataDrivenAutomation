package TestingIM;

import java.io.FileInputStream;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Iterator;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Resources {
    public static String dataSheetpath() {
	String sourcePath = System.getProperty("user.dir") + "\\InputFiles\\PDPTesting.xlsx";
	return sourcePath;
    }

    public static String getData(int row, int cell) throws IOException {
	String path = dataSheetpath();
	FileInputStream fis = null;
	try {
	    fis = new FileInputStream(path);
	} catch (FileNotFoundException e) {
	    e.printStackTrace();
	}
	try (XSSFWorkbook wb = new XSSFWorkbook(fis)) {
	    XSSFSheet sh = wb.getSheetAt(0);
	    XSSFRow r = sh.getRow(row);
	    String c = r.getCell(cell).toString();
	    return c;
	}
    }

    public static boolean setData(int row, int cell, String value) throws IOException {
	try {
	    String path = dataSheetpath();
	    FileInputStream fis = new FileInputStream(path);
	    XSSFWorkbook wb = new XSSFWorkbook(fis);
	    XSSFSheet sh = wb.getSheetAt(0);
	    XSSFRow ro = sh.getRow(row);
	    XSSFCell ce = ro.getCell(cell);
	    ce.setCellValue(value);
	    FileOutputStream fos = new FileOutputStream(path);
	    wb.write(fos);
	    wb.close();
	    return true;
	} catch (Exception e) {
	    System.out.println("Unable to set the data for the cell");
	    return false;
	}
    }

    public static boolean setDataByColumnName(int row, String ColumnName, String setvalue) throws IOException {
	String path = dataSheetpath();
	FileInputStream fis = new FileInputStream(path);
	XSSFWorkbook wb = new XSSFWorkbook(fis);
	XSSFSheet sh = wb.getSheetAt(0);
	Iterator<Row> rows = sh.iterator();
	Row topRow = rows.next();
	Iterator<Cell> cells = topRow.iterator();
	int k = 0;
	int column = 0;
	while (cells.hasNext()) {
	    Cell current = cells.next();
	    if (ColumnName.equalsIgnoreCase(current.getStringCellValue())) {
		column = k;
	    }
	    k++;
	}
	XSSFRow currentRow = sh.getRow(row);
	XSSFCell currentCell = currentRow.getCell(column);
	currentCell.setCellValue(setvalue);
	FileOutputStream fos = new FileOutputStream(path);
	wb.write(fos);
	wb.close();
	return true;
    }

    public static boolean setDataByColumnName1(int row, String columnName, String setValue) throws IOException {
	String path = dataSheetpath();
	FileInputStream fis = new FileInputStream(path);
	XSSFWorkbook wb = new XSSFWorkbook(fis);
	XSSFSheet sh = wb.getSheetAt(0);
	Iterator<Row> rows = sh.iterator();
	Row topRow = rows.next();
	int column = -1;

	for (Cell cell : topRow) {
	    if (columnName.equalsIgnoreCase(cell.getStringCellValue())) {
		column = cell.getColumnIndex();
		break;
	    }
	}

	if (column != -1) {
	    XSSFRow currentRow = sh.getRow(row);
	    XSSFCell currentCell = currentRow.createCell(column);
	    currentCell.setCellValue(setValue);

	    FileOutputStream fos = new FileOutputStream(path);
	    wb.write(fos);
	    wb.close();
	    fis.close();
	    fos.close();
	    return true;
	} else {
	    System.out.println("Column '" + columnName + "' not found.");
	    wb.close();
	    fis.close();
	    return false;
	}
    }

    public static String getDataOfColumn(int row, String columnName) throws IOException {
	String path = dataSheetpath();
	FileInputStream fis = new FileInputStream(path);
	String currentValue = null;

	try (XSSFWorkbook wb = new XSSFWorkbook(fis)) {
	    XSSFSheet sh = wb.getSheetAt(0);
	    Iterator<Row> rows = sh.iterator();
	    Row topRow = rows.next();
	    Iterator<Cell> cells = topRow.iterator();
	    int k = 0;
	    int column = 0;
	    while (cells.hasNext()) {
		Cell curCell = cells.next();
		if (curCell.getStringCellValue().equalsIgnoreCase(columnName)) {
		    column = k;
		}
		k++;
	    }

	    for (int i = 0; i <= sh.getLastRowNum(); i++) {
		XSSFRow currentrow = sh.getRow(row);
		XSSFCell currentCell = currentrow.getCell(column);
		currentValue = currentCell.toString();
	    }
	    return currentValue;
	}

    }
}
