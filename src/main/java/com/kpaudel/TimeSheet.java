package com.kpaudel;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFFormulaEvaluator;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

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
    private static final String TEMPLATE_FILE = "Template.xlsx";
    private static final String DATE_FORMATTER = "yyyy-MM-dd";
    private static final String MONTH_FORMATTER = "dd/MM/yyyy";
    private static final int START_ROW = 10;
    private int month;
    private int year;
    private DateFormat dateFormat;

    public TimeSheet(int year, int month) {
        this.year = year;
        this.month = month;
        this.dateFormat = DateFormat.getDateInstance(DateFormat.FULL);

    }

    public void create() throws IOException, InvalidFormatException {
        String fileName = "Taetigkeitsnachweis_" + this.month + ".xlsx";
        InputStream inputStream = new FileInputStream(TEMPLATE_FILE);
        Workbook workbook = WorkbookFactory.create(inputStream);
        inputStream.close();

        Sheet sheet = workbook.getSheetAt(0);


        Calendar calendar = Calendar.getInstance();
        calendar.set(Calendar.MONTH, this.month);
        calendar.set(Calendar.YEAR, this.year);
        YearMonth yearMonth = YearMonth.of(calendar.get(Calendar.YEAR), calendar.get(Calendar.MONTH));
        String monthYear=new SimpleDateFormat(MONTH_FORMATTER).format(calendar.getTime());
        sheet.getRow(6).getCell(2).setCellValue(monthYear);
        //Navigate the rows
        int day = 0;
        for (int i = START_ROW; i < START_ROW + yearMonth.lengthOfMonth() || day < yearMonth.lengthOfMonth(); i++, ++day) {
            Row row = sheet.getRow(i);
            calendar.set(Calendar.DAY_OF_MONTH, day + 1);
            row.getCell(1).setCellValue(this.dateFormat.format(calendar.getTime()));
            //Holidays
            if (calendar.get(Calendar.DAY_OF_WEEK) == Calendar.SATURDAY || calendar.get(Calendar.DAY_OF_WEEK) == Calendar.SUNDAY) {
                for (int j = 2; j <= 6; j++) {
                    row.getCell(j).setCellValue("");
                }
            }
            //Other days
            else {
                for (int j = 2; j <= 6; j++) {
                    //provide value for 2 4 5 6
                    //row.getCell(j).setCellValue(row.getCell(j).getStringCellValue());
                    if (j == 2 || j == 3) {
                        CellStyle cellStyle = row.getCell(j).getCellStyle();
                    } else {
                        CellStyle cellStyle = row.getCell(j).getCellStyle();
                    }
                }
                row.getCell(2).setCellValue("kampagne");
                String durationFormula = String.format("F%s-E%s", i + 1, i + 1);
                row.getCell(3).setCellFormula(durationFormula);
                row.getCell(4).setCellValue("08:30");
                row.getCell(5).setCellValue("17:00");
                row.getCell(6).setCellValue("00:30");
                String hoursFormula = String.format("(IF(E%s<>\"\",D%s-G%s,\"0\")*24)", i + 1, i + 1, i + 1);
                row.getCell(7).setCellFormula(hoursFormula);
                String workValueFormula = String.format("IF(H%s>=8,1,IF(H%s<8,H%s/8,0))", i + 1, i + 1, i + 1);
                row.getCell(8).setCellFormula(workValueFormula);
                //row.getCell(8).getCellStyle().setFillBackgroundColor(IndexedColors.RED.getIndex());
                //row.getCell(8).getCellStyle().setFillForegroundColor(IndexedColors.RED.getIndex());
                row.getCell(8).getCellStyle().setFillPattern(FillPatternType.SOLID_FOREGROUND);
            }
        }
        for (int i = START_ROW + yearMonth.lengthOfMonth() - 1; i < START_ROW + 31; i++) {
            sheet.removeRow(sheet.getRow(i));
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
