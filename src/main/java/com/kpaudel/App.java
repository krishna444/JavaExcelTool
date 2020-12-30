package com.kpaudel;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

import java.io.FileNotFoundException;
import java.io.IOException;
import java.time.Month;

/**
 * Start from here
 */
public class App {
   public static void main(String[] args) throws FileNotFoundException, IOException, InvalidFormatException {
      TimeSheet timeSheet = new TimeSheet(2020, Month.DECEMBER);
      timeSheet.create();
        /*List<String> items=ExcelUtils.getColumnItems("/home/krishna/Desktop/1und1/1und1/VTRACC/DOTASK-1837/TarifIDsFuerQuickWinV2.xlsx",0,1);

        BufferedWriter writer=new BufferedWriter(new FileWriter("/home/krishna/Desktop/1und1/1und1/VTRACC/DOTASK-1837/TarifIDsFuerQuickWinV2.txt"));
        int elementPerLine=7;
        int count=0;
        for(String item: items){
            writer.append(item);
            writer.append(",");
            count++;
            if(count%elementPerLine==0) {
                writer.append("\n");
                count = 0;
            }
        }
        writer.close();*/
   }
}
