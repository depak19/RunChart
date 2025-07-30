package com.main.excel.dao;

import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.SQLException;

import com.main.excel.app.Configurations;
import com.main.excel.app.Constants;

public class DBConnection {
		
	static Connection con=null;
	static Connection conA=null;
	static String CONFIG_USER = Configurations.getInstance().getProperty(Constants.CONFIG_USER);
	static String ATOMIC_USER = Configurations.getInstance().getProperty(Constants.ATOMIC_USER);
	public static Connection getJDBCConnection(String URL,String USER,String PASS) throws SQLException {
		
		if (USER.equals(CONFIG_USER) && con != null) 
			return con;
		else if (USER.equals(ATOMIC_USER) && conA != null) 
			return conA;
		else {
			
			if(USER.equals(CONFIG_USER)){
				try {
					con = DriverManager.getConnection(URL, USER, PASS);
					//System.out.println("Creating Connection For User: "+USER);
					
				}
				catch (SQLException e) {
					System.out.println(USER+"Connection Failed! Check output console");
					e.printStackTrace();
					return con;
		
				}
				return con;
			}
			else if(USER.equals(ATOMIC_USER)){
				try {
					conA = DriverManager.getConnection(URL, USER, PASS);
					//System.out.println("Creating Connection For User: "+USER);
					
				}
				catch (SQLException e) {
					System.out.println(USER+"Connection Failed! Check output console");
					e.printStackTrace();
					return conA;
		
				}
				return conA;
			}
		}
		return null;
	}
}
