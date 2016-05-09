package com.visa.test.util;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import org.apache.commons.io.IOUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

/**
 * A utility class to read rows of excel files.
 */
public class ExcelReader {

    /**
     * In memory representation of whole excel file.
     */
    private Workbook workBook;

    /**
     * Logging object
     */
    private final Logger logger = LoggerFactory.getLogger(ExcelReader.class);

    /**
     * Use this constructor when a file that is available in the classpath is to be read by the ExcelDataProvider for
     * supporting Data Driven Tests.
     *
     * @param filePath
     *            location of the excel file to be read.
     * @throws IOException
     *             If the file cannot be located, or cannot read by the method.
     */
    public ExcelReader(String filePath) throws Exception {

        if ( StringUtils.isBlank(filePath) ) {
            throw new IllegalArgumentException("filePath cannot be null/empty");
        }

        try {
            workBook = WorkbookFactory.create(new File(filePath));
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    /**
     * Get all excel rows from a specified sheet.
     *
     * @param sheetName
     *             A String that represents the Sheet name
     * @param heading
     *            If true, will return all rows along with the heading row. If false, will return all rows
     *            except the heading row.
     * @return rows that are read.
     */
    public List<Row> getAllExcelRows(String sheetName, boolean heading) {
        Sheet sheet = fetchSheet(sheetName);
        int numRows = sheet.getPhysicalNumberOfRows();
        List<Row> rows = new ArrayList<Row>();
        int currentRow = 1;
        if (heading) {
            currentRow = 0;
        }
        int rowCount = 1;
        while (currentRow <= numRows) {
            Row row = sheet.getRow(currentRow);
            if (row != null) {
                Cell cell = row.getCell(0);
                if (cell != null) {
                    // Did the user mark the current row to be excluded by adding a # ?
                    if (!cell.toString().contains("#")) {
                        rows.add(row);
                    }
                    rowCount = rowCount + 1;
                }
            }
            currentRow = currentRow + 1;
        }
        return rows;
    }

    /**
     * A utility method, which returns {@link Sheet} for a given sheet name.
     *
     * @param sheetName - A string that represents a valid sheet name.
     * @return - An object of type {@link Sheet}
     */
    protected Sheet fetchSheet(String sheetName) {
        Sheet sheet = workBook.getSheet(sheetName);
        if (sheet == null) {
            IllegalArgumentException e = new IllegalArgumentException("Sheet '" + sheetName + "' is not found.");
            throw e;
        }
        return sheet;
    }

    /**
     * A Utility method to find if a sheet exists in the workbook
     *
     * @param sheetName - A String that represents the Sheet name
     * @return true if the sheet exists, false otherwise
     */
    public boolean sheetExists(String sheetName) {
        return (workBook.getSheet(sheetName) != null);
    }

    /**
     * Using the specified rowIndex to search for the row from the specified Excel sheet, then return the row contents
     * in a list of string format.
     *
     * @param rowIndex
     *            - The row number from the excel sheet that is to be read. For e.g., if you wanted to read the 2nd row
     *            (which is where your data exists) in your excel sheet, the value for index would be 1. <b>This method
     *            assumes that your excel sheet would have a header which it would EXCLUDE.</b> When specifying index
     *            value always remember to ignore the header, since this method will look for a particular row ignoring
     *            the header row.
     * @param size
     *            - The number of columns to read, including empty and blank column.
     * @return List<String> String array contains the read data.
     */
    public List<String> getRowContents(String sheetName, int rowIndex, int size) {
        Sheet sheet = fetchSheet(sheetName);

        int actualExcelRow = rowIndex - 1;
        Row row = sheet.getRow(actualExcelRow);

        List<String> rowData = getRowContents(row, size);

        return rowData;
    }

    /**
     * Return the row contents of the specified row in a list of string format.
     *
     * @param size - The number of columns to read, including empty and blank column.
     * @return List<String> String array contains the read data.
     */
    public List<String> getRowContents(Row row, int size) {
        List<String> rowData = new ArrayList<String>();
        if (row != null) {
            for (int i = 1; i <= size; i++) {
                String data = null;
                if (row.getCell(i) != null) {
                    data = row.getCell(i).toString();
                }
                rowData.add(data);
            }
        }
        return rowData;
    }

    /**
     * Search for the input key from the specified sheet name and return the index position of the row that contained
     * the key
     *
     * @param sheetName
     *            - A String that represents the Sheet name from which data is to be read
     * @param key
     *            - A String that represents the key for the row for which search is being done.
     * @return - An int that represents the row number for which the key matches. Returns -1 if the search did not yield
     *         any results.
     *
     */
    public int getRowIndex(String sheetName, String key) {
        int index = -1;
        Sheet sheet = fetchSheet(sheetName);

        int rowCount = sheet.getPhysicalNumberOfRows();
        for (int i = 0; i < rowCount; i++) {
            Row row = sheet.getRow(i);
            if (row == null) {
                continue;
            }
            String cellValue = row.getCell(0).toString();
            if ((key.compareTo(cellValue) == 0) && (!cellValue.contains("#"))) {
                index = i;
                break;
            }
        }
        return index;
    }

    /**
     *
     * @param sheetName -  A String that represents the Sheet name from which data is to be read
     * @param rowNumber - The row number from the excel sheet that is to be read
     * @return - Single Excel row that was read
     */
    public Row getAbsoluteSingeExcelRow(String sheetName, int rowNumber) {
        Sheet sheet = fetchSheet(sheetName);
        Row row = sheet.getRow(rowNumber);
        return row;
    }

}