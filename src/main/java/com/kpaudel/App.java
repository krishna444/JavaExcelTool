package com.kpaudel;

import java.io.*;
import java.time.Month;
import java.util.*;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * Hello world!
 *
 */
public class App 
{
    public static void main( String[] args ) throws FileNotFoundException, IOException, InvalidFormatException {
        TimeSheet timeSheet=new TimeSheet(2020, Month.JANUARY);
        timeSheet.create();
        //List<String> items=ExcelUtils.getColumnItems("/home/krishna/Desktop/1und1/Tools/VTRACC.2757_Nutzerkonzept_2019-10-25_V.1.3.xlsx",0,1);

        //BufferedWriter writer=new BufferedWriter(new FileWriter("/home/krishna/Desktop/1und1/Tools/VTRACC.2757_Nutzerkonzept_2019-10-25_V.1.3_list.txt"));
        //for(String item: items){
        //    writer.append(item);
        //    writer.append(",");
        //}
        //writer.close();

    }
}
