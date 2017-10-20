package com.ttgint.db;

import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;

public class OraDBConnection implements IDBConnection {

    Connection con = null;

    private final String url;
    private final String userName;
    private final String passWord;

    public OraDBConnection(String url, String userName, String passWord) {
        this.url = url;
        this.userName = userName;
        this.passWord = passWord;
    }

    @Override
    public Connection getConnection() throws SQLException, InterruptedException {
        if (con == null) {
            try {
                con = DriverManager.getConnection(url, userName, passWord);
            } catch (SQLException ex) {
                System.out.println("Error: No DB connection");
            }
        }
        return con;
    }

    @Override
    public void Execute(String sql) throws Exception {
        Statement stmt = null;
        try {
            stmt = con.createStatement();
            stmt.execute(sql);

        } catch (SQLException ex) {
        } finally {
            if (stmt != null) {
                try {
                    stmt.close();
                } catch (SQLException ex) {
                }
            }
        }
    }

    public void reNewConnection() throws Exception {
        System.out.println("Connection is renewing");
        try {
            getConnection();
            if (isConnectionAlive()) {
                System.out.println("Connection is renewed");
            }
        } catch (SQLException | InterruptedException ex) {
            System.out.println("Connection renewing is failed");
            ex.printStackTrace();
        }
    }

    @Override
    public void closeConnection() {
        try {
            con.close();
        } catch (SQLException ex) {
            System.out.println(ex.toString());
        }
    }

    //to check is connection alive , this is only certain way to test db connection
    @Override
    public boolean isConnectionAlive() {
        try {
            ResultSet rss
                    = con.createStatement()
                    .executeQuery("select sysdate from dual");
            rss.next();
            rss.close();
            return true;
        } catch (Exception ex) {

            return false;
        }
    }

    public void refreshConnection() throws SQLException, InterruptedException {
        if (con != null && con.isClosed() == false) {
            closeConnection();
        }
        con = null;
        getConnection();
    }

    @Override
    public ResultSet ExecuteQuery(String sql) {
        try {
            return con.createStatement(ResultSet.TYPE_SCROLL_INSENSITIVE, ResultSet.CONCUR_READ_ONLY).executeQuery(sql);
        } catch (SQLException ex) {
            ex.printStackTrace();
        }
        return null;
    }
}
