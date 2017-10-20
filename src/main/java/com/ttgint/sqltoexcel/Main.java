/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.ttgint.sqltoexcel;

import com.ttgint.db.OraDBConnection;
import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileReader;
import java.io.IOException;
import java.io.InputStream;
import java.sql.ResultSet;
import java.sql.ResultSetMetaData;
import java.util.ArrayList;
import java.util.List;
import java.util.Properties;
import java.util.TreeMap;

/**
 *
 * @author erdigurbuz
 */
public class Main {

    static OraDBConnection db;

    public static void main(String[] args) throws IOException {
        String reportFileName = "Result.xlsx";

        try {
            connectDB();

            TreeMap<String, ArrayList<String>> sqlSheetMap = new TreeMap<>();
            sqlSheetMap = getSqlBySheet();
            for (String sheetName : sqlSheetMap.keySet()) {
                for (String sql : sqlSheetMap.get(sheetName)) {
                    sqlToExcel(reportFileName, sheetName, sql.replace(";", ""));
                }
            }

            closeDB();
        } catch (IOException ex) {
            ex.printStackTrace();
        }
    }

    public static void sqlToExcel(String excelName, String reportName, String query) {
        try {
            ExcelFileUtil util = new ExcelFileUtil(excelName, reportName);
            ResultSet res = db.ExecuteQuery(query);
            ResultSetMetaData rsmd = res.getMetaData();

            //Write header
            List<String> header = new ArrayList<>();
            for (int index = 1; index <= rsmd.getColumnCount(); index++) {
                header.add(rsmd.getColumnName(index));
            }
            util.writeLine(header);

            //Write data
            while (res.next()) {
                List<String> row = new ArrayList<>();
                for (int index = 1; index <= rsmd.getColumnCount(); index++) {
                    row.add(res.getString(index));
                }
                util.writeLine(row);
            }

            res.close();

            //Create excel
            util.createWorkBook();
        } catch (Exception ex) {
            ex.printStackTrace();
        }
    }

    public static TreeMap getSqlBySheet() throws IOException {
        BufferedReader buf = new BufferedReader(new FileReader(new File("sql.properties")));
        String line;

        TreeMap<String, ArrayList<String>> map = new TreeMap<>();
        String tmpSheetName = "";
        String sql = "";
        while ((line = buf.readLine()) != null) {
            if (line.length() < 2) {
            }
            if (line.startsWith("#")) {
                String sheetName = line.split("#")[1].trim();
                map.put(sheetName, new ArrayList<>());
                tmpSheetName = sheetName;
            } else if (line.contains(";")) {
                sql = line.trim();
                ArrayList<String> ls = map.get(tmpSheetName);
                ls.add(sql);
                map.put(tmpSheetName, ls);
            } else {
                sql += " " + line;
            }
        }
        buf.close();

        return map;
    }

    public static void connectDB() throws FileNotFoundException, IOException {
        Properties prop = new Properties();
        InputStream cfg = new FileInputStream("settings.properties");
        prop.load(cfg);

        try {
            System.out.println("Connecting to DB...");
            //DBConnection
            if (db == null || !db.isConnectionAlive()) {
                try {
                    String url = "jdbc:oracle:thin:@" + prop.getProperty("DB_HOSTNAME") + ":" + prop.getProperty("DB_PORT") + "/" + prop.getProperty("DB_SID");
                    db = new OraDBConnection(url, prop.getProperty("DB_USERNAME"), prop.getProperty("DB_PASSWORD"));
                    db.getConnection();
                    if (db.isConnectionAlive()) {
                        System.out.println("Connected...");
                    } else {
                        db = null;
                    }
                } catch (Exception ex) {
                    ex.printStackTrace();
                }
            }
        } catch (Exception ex) {
            ex.printStackTrace();
        }

    }

    public static void closeDB() {
        if (db != null && db.isConnectionAlive()) {
            db.closeConnection();
        }
    }

}
