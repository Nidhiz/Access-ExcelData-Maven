package com.excel;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.NumberToTextConverter;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelDemo {
	
	public ArrayList<String> getData(String testcaseName) throws IOException {

		FileInputStream fis = new FileInputStream("D:\\ExcelFiles\\Data Refresh Sample Data.xlsx");

		XSSFWorkbook workbook = new XSSFWorkbook(fis);

		int sheets = workbook.getNumberOfSheets();
		ArrayList<String> a = new ArrayList<String>();

		for (int i = 0; i < sheets; i++) {

			if (workbook.getSheetName(i).equalsIgnoreCase("Sheet1")) {
				// identify column by scanning entire first row
				XSSFSheet xsheet = workbook.getSheetAt(i);

				Iterator<Row> rows = xsheet.iterator();// sheet is collection of rows
				Row firstRow = rows.next();
				Iterator<Cell> cells = firstRow.cellIterator();// Row is collection of cells

				int k = 0, column = 0;

				while (cells.hasNext()) {
					Cell value = cells.next();
					if (value.getStringCellValue().equalsIgnoreCase("Ship Mode")) {
						column = k;
					}
					k++;
				}
				// System.out.println(column);

				//// once column is identified then scan entire column to identify
				//// testcase row
				while (rows.hasNext()) {

					Row r = rows.next();

					if (r.getCell(column).getStringCellValue().equalsIgnoreCase(testcaseName)) {

						// after you grab row = pull all the data of that row and feed into test

						Iterator<Cell> cv = r.cellIterator();
						while (cv.hasNext()) {
							Cell c = cv.next();
							// to get numeric value from excel
							if (c.getCellTypeEnum() == CellType.STRING) {
								a.add(c.getStringCellValue());
							} else {
								// convert back to string this numeric value
								a.add(NumberToTextConverter.toText(c.getNumericCellValue()));

							}
						}
					}

				}

			}
		}
		return a;
	}

}
