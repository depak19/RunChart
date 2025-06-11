package com.main.excel.util;

import java.io.IOException;
import java.security.NoSuchAlgorithmException;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashSet;
import java.util.List;
import java.util.zip.DataFormatException;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import com.main.excel.app.Configurations;
import com.main.excel.app.Constants;
import com.main.excel.dao.MetadataDao;
import com.main.excel.model.XmlData;
import com.main.excel.service.ReadT2TXml;
import com.trivadis.unwrapper.util.Unwrapper;

public class AppUtil {
	
	static String SOURCE_DIR = Configurations.getInstance().getProperty(Constants.SOURCE_DIR);
	static String SOURCE_NAME = Configurations.getInstance().getProperty(Constants.SOURCE_NAME);
	
	public static List<String> removeDuplicates(List<String> list) {
		ArrayList<String> result = new ArrayList<String>();
		HashSet<String> set = new HashSet<String>();
		for (String item : list) {
		    if (!set.contains(item)) {
		    	result.add(item);
		    	set.add(item);
		    }
		}
		return result;
	}
	
	public static void fillRow(HSSFRow row,HSSFCellStyle style,int j,String str){
		HSSFCell cell = null;
		cell = row.createCell(j);
		cell.setCellValue(str);
		cell.setCellStyle(style);
	}
	
	public static String getSheetName (String str) {
		str = str.replaceAll("\\W", "");
		int len = str.length();
		if (len > 25)
			str = str.substring(0, 25) ;
		else
			str = str.substring(0, len);
		return str;
	}
	public static int getRowSetSize(ResultSet rs1) throws SQLException{
		int count=0;
		while(rs1.next()){
			count++;
		}
		rs1.beforeFirst();
		return count;
	}
	
	public static  XmlData getT2TDetails(HSSFWorkbook hwb,String taskName,String sheetName) throws Exception{
		return ReadT2TXml.readXMLFileCreateXLS(SOURCE_DIR,taskName,SOURCE_NAME);
	}
	
	public static String showTime(){
		DateFormat dateFormat = new SimpleDateFormat("yyyy/MM/dd HH:mm:ss");
		Date date = new Date();
		return dateFormat.format(date);
	}
	
	public static String showDate(){
		DateFormat dateFormat = new SimpleDateFormat("yyyyMMdd");
		Date date = new Date();
		return dateFormat.format(date);
	}
	public static StringBuilder getObjectName(StringBuilder unWrapText,List<String> functionList) throws SQLException, NoSuchAlgorithmException, DataFormatException, IOException{
		MetadataDao metadata = new MetadataDao();
		ResultSet rs = null;
		for (String funName : functionList){
			if (unWrapText.indexOf(funName) != -1){
				//System.out.println("Function Name:"+funName);
				rs = metadata.getDTDetails(funName);
			}	
		}
		StringBuilder wrapText= new StringBuilder();
		if(rs!=null){
			while (rs.next()) {
				wrapText.append(rs.getString("TEXT"));
			}
			//System.out.println("wrapText:::"+wrapText.toString());
		}
		StringBuilder unWrapText1 = new StringBuilder();
		if (wrapText.indexOf("wrapped")!=-1){
			unWrapText1.append(Unwrapper.unwrap(wrapText.toString()));
		}
		else 
			unWrapText1.append(wrapText);
		return unWrapText1;
	}
	
	public static  List<String> getSrcTargetHier(ResultSet rs) throws Exception{
		List <String> srcTargetHier = new ArrayList<String>();
		int count=0;
		while (rs.next()) {
			if(!rs.getString("OBJECT_TYPE_CODE").equals("BP"))
				srcTargetHier.add(rs.getString("DESC1"));
			else if(count==0) {
				count++;
				srcTargetHier.add("Target BP");
			}
		}
		return srcTargetHier;
	}
	public static String getMergeCell(char ch,int i,int len){
		String str = ch+""+i+":"+(char)((int)ch+len)+i;
		//System.out.println("Merge Cell::"+str);
		return str;
	}
}
