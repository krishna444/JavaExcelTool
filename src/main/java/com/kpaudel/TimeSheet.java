package com.kpaudel;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFFormulaEvaluator;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import sun.util.calendar.CalendarUtils;

import java.io.*;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.time.YearMonth;
import java.util.*;

/**
 * @author krishna<br>
 * Date 25.10.19
 */
public class TimeSheet {
    private static String TEMPLATE_FILE = "Template.xlsx";
    private static String DATE_FORMATTER = "yyyy-MM-dd";
    private static int START_ROW = 10;
    private int month;
    private int year;
    private DateFormat dateFormat;

    public TimeSheet(int year, int month) {
        this.year=year;
        this.month = month;
        this.dateFormat = DateFormat.getDateInstance(DateFormat.FULL);

    }

    public void create() throws IOException, InvalidFormatException {
        String fileName = "Taetigkeitsnachweis_" + this.month + ".xlsx";
        InputStream inputStream = new FileInputStream(TEMPLATE_FILE);
        Workbook workbook = WorkbookFactory.create(inputStream);
        inputStream.close();

        Sheet sheet = workbook.getSheetAt(0);
        //Navigate the rows
        Calendar calendar = Calendar.getInstance();
        calendar.set(Calendar.MONTH,this.month);
        calendar.set(Calendar.YEAR,this.year);
        YearMonth yearMonth = YearMonth.of(calendar.get(Calendar.YEAR), calendar.get(Calendar.MONTH));
        int day = 0;
        for (int i = START_ROW; i <= START_ROW + yearMonth.lengthOfMonth() || day<=yearMonth.lengthOfMonth(); i++,++day) {
            Row row = sheet.getRow(i);
            calendar.set(Calendar.DAY_OF_MONTH, day);
            row.getCell(1).setCellValue(this.dateFormat.format(calendar.getTime()));
            row.getCell(1).setCellStyle(row.getCell(1).getCellStyle());
        }
        XSSFFormulaEvaluator.evaluateAllFormulaCells(workbook);

        OutputStream outputStream = new FileOutputStream(fileName);
        workbook.write(outputStream);
        outputStream.close();
        workbook.close();
/*
        Sheet sheet = workbook.createSheet("TimeSheet1");
        sheet.setColumnWidth(0, 400);
        sheet.setColumnWidth(1, 4000);

        Row row = sheet.createRow(1);
        Cell cell = row.createCell(1);
        CellStyle headingTitleStyle = workbook.createCellStyle();
        Font font=workbook.createFont();
        font.setColor(HSSFColor.BLACK.index);
        headingTitleStyle.setFont(font);
        headingTitleStyle.setBorderLeft(BorderStyle.THICK);
        headingTitleStyle.setBorderTop(BorderStyle.THICK);
        headingTitleStyle.setBorderRight(BorderStyle.THICK);
        headingTitleStyle.setBorderBottom(BorderStyle.THICK);
        headingTitleStyle.setFillForegroundColor(IndexedColors.SKY_BLUE.getIndex());
        headingTitleStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        //headingTitleStyle.setFillPattern(CellStyle.BIG_SPOTS);
        headingTitleStyle.setWrapText(true);
        cell.setCellValue("Krishna");
        cell.setCellStyle(headingTitleStyle);

        File currDir = new File(".");
        String path = currDir.getAbsolutePath();
        String fileLocation = path.substring(0, path.length() - 1) + fileName;

        FileOutputStream outputStream=new FileOutputStream(fileLocation);
        workbook.write(outputStream);
        workbook.close();*/

    }

    public void read() throws IOException {
        String PATH = "/home/krishna/Desktop/1und1/TestWorkspace/JavaTools/exceltool/Template.xlsx";
        FileInputStream file = new FileInputStream(new File(PATH));
        Workbook workbook = new XSSFWorkbook(file);
        Sheet sheet = workbook.getSheetAt(0);
        Map<Integer, List<String>> data = new HashMap();
        int i = 0;
        for (Row row : sheet) {
            data.put(i, new ArrayList<String>());
            for (Cell cell : row) {
                switch (cell.getCellTypeEnum()) {
                    case STRING:
                        System.out.println(cell.getStringCellValue());
                }
            }
            i++;
        }
    }
}
