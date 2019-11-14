package com.kpaudel;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

/**
 * Utility class for Excel
 * @author krishna<br>
 * Date 14.11.19
 */
public class ExcelUtils {
    /**
     * Gets the items of a colums as a list.
     * @param fileName File Name
     * @param columnNumber Column number
     * @param startRow Starting row(normally 0, but we could discard heading row(s)
     * @return List of items of a column
     * @throws IOException
     * @throws InvalidFormatException
     */
    public static List<String> getColumnItems(String fileName, int columnNumber, int startRow) throws IOException, InvalidFormatException {
        List<String> itemList = new ArrayList<>();
        InputStream inputStream = new FileInputStream(fileName);
        XSSFWorkbook workbook = (XSSFWorkbook) WorkbookFactory.create(inputStream);
        inputStream.close();

        XSSFSheet sheet = workbook.getSheetAt(0);
        workbook.close();

        int rowIndex = startRow;
        boolean stop = false;
        while (sheet.getRow(rowIndex) != null) {
            System.out.println(rowIndex);
            XSSFCell cell = sheet.getRow(rowIndex).getCell(columnNumber);

            itemList.add(cell.getRawValue());
            rowIndex++;
        }
        return itemList;

    }
}
