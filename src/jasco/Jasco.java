/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package jasco;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.logging.Level;
import java.util.logging.Logger;
import net.arnx.jsonic.JSON;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;


public class Jasco {
    //This tool is WIP

    public static void main(String[] args) {
        try {
            //TODO Read from JSON
            //TODO Read from args

            int x = 2;
            int y = 24;
            int width = 19;
            int height = 29;
            final String sheetName = "SheetName";
            final String xlsFile = "data/test.xls";
            
            new Jasco().exec(xlsFile, sheetName, x, y, width, height);
        } catch (FileNotFoundException ex) {
            Logger.getLogger(Jasco.class.getName()).log(Level.SEVERE, null, ex);
        }
    }

    private void exec(String xlsFile, String sheetName,int x,int y,int width,int height) throws FileNotFoundException {
        InputStream in = new FileInputStream(xlsFile);
        Workbook wb;
        try {
            wb = WorkbookFactory.create(in);
            Sheet sheet1 = wb.getSheet(sheetName);

            List<List<String>> rows = convertSheetToArrayList(y, height, x, width, sheet1);

            String encode = JSON.encode(rows);
            
            //TODO output File
            System.out.println(encode);

        } catch (IOException | InvalidFormatException ex) {
            Logger.getLogger(Jasco.class.getName()).log(Level.SEVERE, null, ex);
        }

    }

    private List<List<String>> convertSheetToArrayList(int y, int height, int x, int width, Sheet sheet1) {
        List<List<String>> rows = new ArrayList<>();
        for (int row = y; row < y + height; row++) {
            ArrayList<String> rowData = new ArrayList<>();
            for (int col = x; col < x + width; col++) {
                Cell cell = sheet1.getRow(row).getCell(col);
                switch (cell.getCellType()) {
                    case (Cell.CELL_TYPE_STRING):
                        RichTextString str = cell.getRichStringCellValue();
                        rowData.add(str.toString());
                        break;
                    case (Cell.CELL_TYPE_NUMERIC):
                        rowData.add("" + cell.getNumericCellValue());
                        break;
                    case (Cell.CELL_TYPE_BOOLEAN):
                        rowData.add("" + cell.getBooleanCellValue());
                        break;
                    case (Cell.CELL_TYPE_FORMULA):
                        switch (cell.getCachedFormulaResultType()) {
                            case Cell.CELL_TYPE_NUMERIC:
                                rowData.add("" + cell.getNumericCellValue());
                                break;
                            case Cell.CELL_TYPE_STRING:
                                rowData.add("" + cell.getRichStringCellValue());
                                break;
                        }
                        break;
                    case (Cell.CELL_TYPE_BLANK):
                        rowData.add("");
                        break;
                    default:
                        System.out.println("unknown cell type" + cell.getCellType());
                }
                
                rows.add(rowData);
            }
        }
        return rows;
    }
}
