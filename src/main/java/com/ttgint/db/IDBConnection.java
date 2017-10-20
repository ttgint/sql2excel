
package com.ttgint.db;

import java.sql.Connection;
import java.sql.ResultSet;

public interface IDBConnection {
    
    public Connection getConnection() throws Exception;
    public void Execute(String sql) throws  Exception;
    public ResultSet ExecuteQuery(String sql) throws Exception;
    public void closeConnection();
    public boolean isConnectionAlive() throws Exception;
    
}
