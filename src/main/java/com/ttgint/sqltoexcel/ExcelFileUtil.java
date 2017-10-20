/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.ttgint.sqltoexcel;

/**
 *
 * @author erdigurbuz
 */
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelFileUtil {

    Workbook wb;
    Sheet sheet;
    public String filename;
    int cellrowid = 0;

    public ExcelFileUtil(String filename, String sheetName) throws IOException, InvalidFormatException {
        this.filename = filename;
        File file = new File(filename);

        if (file.exists()) {
            try {
                POIFSFileSystem fs = new POIFSFileSystem(new FileInputStream(file));
                wb = new HSSFWorkbook(fs);
            } catch (Exception ex) {
                wb = new XSSFWorkbook(new FileInputStream(file));
            }

            try {
                sheet = wb.getSheet(sheetName);
                cellrowid = sheet.getPhysicalNumberOfRows();
            } catch (Exception ex) {
                sheet = wb.createSheet();
                wb.setSheetName(wb.getSheetIndex(sheet), sheetName);
                cellrowid = sheet.getPhysicalNumberOfRows();
            }

        } else {
            wb = new HSSFWorkbook();
            sheet = wb.createSheet();
            wb.setSheetName(wb.getSheetIndex(sheet), sheetName);
        }

    }

    public void writeLine(List<String> inputStr) {
        Row row = sheet.createRow(cellrowid);
        cellrowid = cellrowid + 1;
        for (int i = 0; i < inputStr.size(); i++) {
            String s = inputStr.get(i);
            Cell cell = row.createCell(i);
            cell.setCellType(HSSFCell.CELL_TYPE_STRING);
            cell.setCellValue(s);
        }
    }

    public void writeLine(String... inputStr) {
        Row row = sheet.createRow(cellrowid);
        cellrowid = cellrowid + 1;
        for (int i = 0; i < inputStr.length; i++) {
            Cell cell = row.createCell(i);
            cell.setCellType(HSSFCell.CELL_TYPE_STRING);
            cell.setCellValue(inputStr[i]);
        }
    }

    public void createWorkBook() throws IOException {
        FileOutputStream fileOut = new FileOutputStream(this.filename);
        wb.write(fileOut);
        fileOut.close();
    }

}
