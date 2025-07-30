package com.main.excel.service;

import java.sql.ResultSet;
import java.sql.SQLException;
import java.util.ArrayList;
import java.util.List;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.util.CellRangeAddress;

import com.main.excel.app.Constants;
import com.main.excel.model.T2TMapping;
import com.main.excel.model.XmlSrcDestTableMap;
import com.main.excel.util.AppUtil;

public class MetadataDetailSheet {

	public MetadataDetailSheet() {
		super();
	}
	
	public int ruleDetailPopulation (HSSFWorkbook hwb,HSSFSheet sheet,ResultSet rs) throws SQLException {
		HSSFCell cell = null;
		HSSFRow row = null;
		HSSFCellStyle style = null;
		String color = Constants.COLOR3;
		int i=3;
		String[] names = new String[] {"Name","Rule Code","Short Description","Created By","Created Date","Last Modified By","Last Modified Date","Authorized By","Authorized Date","DataSet Code","Type","DataSet"};
		while (rs.next()) {
			int j=1;
			row = sheet.createRow((short) 2);
			cell = row.createCell(1);
			sheet.addMergedRegion(CellRangeAddress.valueOf("B3:C3"));
			cell.setCellValue(rs.getString(2));
			style=MetadataStyleSheet.getCellStyleHeaderH1(hwb);
			cell.setCellStyle(style);
			cell = row.createCell(2);
			cell.setCellStyle(style);
			for(String name :names) {
				if(color.equals(Constants.COLOR3)) {
					style=MetadataStyleSheet.getCellStyleGray(hwb,0);
					color = Constants.COLOR2;
				}
				else {
					style=MetadataStyleSheet.getCellStyleBlue(hwb,0);
					color = Constants.COLOR3;
				}
				row = sheet.createRow((short) i);
				cell = row.createCell( 1);
				cell.setCellValue(name);
				cell.setCellStyle(style);
				cell = row.createCell(2);
				if(rs.getString(j)!=null)
					cell.setCellValue(StringUtils.normalizeSpace(rs.getString(j)).replace("\n"," "));
				else
					cell.setCellValue(rs.getString(j));
				cell.setCellStyle(style);
				i++;
				j++;
					
			}
		}
		return i;
	}
	public int ruleHireBPPopulation(HSSFWorkbook hwb,HSSFSheet sheet,ResultSet rs,int i) throws SQLException {
		HSSFCell cell = null;
		HSSFRow row = null;
		HSSFCellStyle style = null;
		String color = Constants.COLOR3;
		String str1="ABC";
		i=i+1;
		while (rs.next()) {
			String mergeCell = null;
			if(!str1.equals(rs.getString("COL1"))) {
			row = sheet.createRow((short) i);
			cell = row.createCell(1);
			i++;
			mergeCell ="B"+i+":C"+i;
			sheet.addMergedRegion(CellRangeAddress.valueOf(mergeCell));
			style=MetadataStyleSheet.getCellStyleHeaderH1(hwb);
			cell.setCellValue(rs.getString("COL1"));
			cell.setCellStyle(style);
			cell = row.createCell(2);
			cell.setCellStyle(style);
			}
			row = sheet.createRow((short) i);
			i++;
			if(color.equals(Constants.COLOR3)) {
				//mergeCell ="B"+i+":C"+i;
				//sheet.addMergedRegion(CellRangeAddress.valueOf(mergeCell));
				style=MetadataStyleSheet.getCellStyleGray(hwb,0);
				color = Constants.COLOR2;
			}
			else {
				//mergeCell ="B"+i+":C"+i;
				//sheet.addMergedRegion(CellRangeAddress.valueOf(mergeCell));
				style=MetadataStyleSheet.getCellStyleBlue(hwb,0);
				color = Constants.COLOR3;
			}
			str1=rs.getString("COL1");
			cell = row.createCell(1);
			cell.setCellValue(rs.getString("CODE1")+"-->"+rs.getString("DESC1"));
			cell.setCellStyle(style);
			cell = row.createCell(2);
			cell.setCellValue(rs.getString("VALUE1"));
			cell.setCellStyle(style);
		}
		return i;
	}
	public List<String> ruleMappingPopulation(HSSFWorkbook hwb,HSSFSheet sheet,ResultSet rs,int i,List<String> arrayList) throws SQLException {
		HSSFCell cell = null;
		HSSFRow row = null;
		HSSFCellStyle style = null;
		String color = Constants.COLOR3;	
		String mergeCell = null;
		List<String> list = new ArrayList<String>();
		int listSize=arrayList.size();
		i++;
		row = sheet.createRow((short) i);
		i++;
		//mergeCell ="B"+i+":C"+i;
		mergeCell=AppUtil.getMergeCell('B', i, listSize-1);
		sheet.addMergedRegion(CellRangeAddress.valueOf(mergeCell));
		style=MetadataStyleSheet.getCellStyleHeaderH1(hwb);
		cell = row.createCell(1);
		cell.setCellValue("Mapping Details");
		cell.setCellStyle(style);
		for(int k=2;k<=listSize;k++){
			cell = row.createCell(k);
			cell.setCellStyle(style);
		}
		row = sheet.createRow((short) i);
		i++;
		mergeCell=AppUtil.getMergeCell('B', i, listSize-2);
		sheet.addMergedRegion(CellRangeAddress.valueOf(mergeCell));
		style=MetadataStyleSheet.getCellStyleHeaderH1(hwb);
		cell = row.createCell(1);
		cell.setCellValue("Source");
		cell.setCellStyle(style);
		for(int k=2;k<=listSize-1;k++){
			cell = row.createCell(k);
			cell.setCellStyle(style);
		}
		cell = row.createCell(listSize);
		cell.setCellValue("Target");
		cell.setCellStyle(style);
		row = sheet.createRow((short) i);
		i++;
		style=MetadataStyleSheet.getCellStyleHeaderH2(hwb);
		for(int j=1;j<listSize;j++) {
			cell = row.createCell(j);
			if(j!=1 || j!=2){
				sheet.setColumnWidth(j,(Constants.cellWidth*256));
			}
			cell.setCellValue(arrayList.get(j-1));
			cell.setCellStyle(style);
			
		}
		cell = row.createCell(listSize);
		sheet.setColumnWidth(listSize,(Constants.cellWidth*256));
		cell.setCellValue(arrayList.get(listSize-1));
		cell.setCellStyle(style);
		while (rs.next()) {
			if(color.equals(Constants.COLOR3)) {
				style=MetadataStyleSheet.getCellStyleGray(hwb,0);
				color = Constants.COLOR2;
			}
			else {
				style=MetadataStyleSheet.getCellStyleBlue(hwb,0);
				color = Constants.COLOR3;
			}
			row = sheet.createRow((short) i);
			for(int j =1;j<listSize;j++) {
				cell = row.createCell(j);
				cell.setCellValue(rs.getString("S"+j));
				cell.setCellStyle(style);
			}
			cell = row.createCell(listSize);
			cell.setCellValue(rs.getString("TARGET_NAME"));
			cell.setCellStyle(style);
			if(rs.getString("METADATA_TYPE").equals("BP")) {
				list.add(rs.getString("TARGET_MEMBER"));
			}
			i++;
			
		}
		list.add(i+"");
		return AppUtil.removeDuplicates(list);
	}
	
	public int ruleBPDetailPopulation (HSSFWorkbook hwb,HSSFSheet sheet,ResultSet rs,int i) throws SQLException {
		HSSFCell cell = null;
		HSSFRow row = null;
		HSSFCellStyle style = null;
		String color = Constants.COLOR3;
		String[] names = new String[] {"Code","Description","Created By","Created Date","Last Modified By","Last Modified Date","Authorized By","Authorized Date","DataSet","Target Measure","Aggregation Function","Expression","Used Measures"};
		while (rs.next()) {
			int j=1;
			row = sheet.createRow((short) i);
			cell = row.createCell(1);
			i++;
			String mergeCell ="B"+i+":C"+i;
			sheet.addMergedRegion(CellRangeAddress.valueOf(mergeCell));
			cell.setCellValue(rs.getString(2));
			style=MetadataStyleSheet.getCellStyleHeaderH1(hwb);
			cell.setCellStyle(style);
			cell = row.createCell(2);
			cell.setCellStyle(style);
			for(String name :names) {
				if(color.equals(Constants.COLOR3)) {
					style=MetadataStyleSheet.getCellStyleGray(hwb,0);
					color = Constants.COLOR2;
				}
				else {
					style=MetadataStyleSheet.getCellStyleBlue(hwb,0);
					color = Constants.COLOR3;
				}
				row = sheet.createRow((short) i);
				cell = row.createCell( 1);
				cell.setCellValue(name);
				cell.setCellStyle(style);
				cell = row.createCell(2);
				cell.setCellValue(rs.getString(j));
				cell.setCellStyle(style);
				i++;
				j++;		
			}
			i++;
			
		}
		return i;
	}
	
	
	public void t2TDetailPopulation(HSSFWorkbook hwb,HSSFSheet sheet,T2TMapping xmlData) {
		int i=5;
		HSSFRow row = sheet.createRow((short) i);
		HSSFCellStyle style = MetadataStyleSheet.getCellStyleHeaderH2(hwb);
		HSSFCellStyle styleLabel = MetadataStyleSheet.getCellStyleHeaderH2(hwb);
		HSSFCellStyle styleValue = MetadataStyleSheet.getCellStyleWhite(hwb,0);
		AppUtil.fillRow(row,styleLabel,1,"ANSI Join");
		
		//System.out.println("Ansi Join length :"+xmlData.getAnsiJoin().length());
		int ansiLength=0;
		int start=0;
		if (xmlData.getAnsiJoin()!=null)
			ansiLength=xmlData.getAnsiJoin().length();
		
		String cellLoc=null;
		while(start < ansiLength) {
		cellLoc="C"+i+":E"+i;
		sheet.addMergedRegion(CellRangeAddress.valueOf(cellLoc));
		int end=((ansiLength-start)>=32000 ? 32000 : ansiLength-start)+start;
		//System.out.println("Start:"+start+" : End :"+end+" CellLoc:"+cellLoc);
		String trimAnsiJoin=xmlData.getAnsiJoin().substring (start, end);
		//System.out.println(" Ans JOin: "+trimAnsiJoin);
		AppUtil.fillRow(row,styleValue,2,trimAnsiJoin);
		AppUtil.fillRow(row,styleValue,3,"");
		AppUtil.fillRow(row,styleValue,4,"");
		start =start+32000;
		i++;
		row = sheet.createRow((short) i);
		}
		cellLoc="C"+i+":E"+i;
		sheet.addMergedRegion(CellRangeAddress.valueOf(cellLoc));
		int k =i+1;
		cellLoc="C"+k+":E"+k;
		//row = sheet.createRow((short) i);
		AppUtil.fillRow(row,styleLabel,1,"Filter");
		sheet.addMergedRegion(CellRangeAddress.valueOf(cellLoc));
		AppUtil.fillRow(row,styleValue,2,xmlData.getFilter());
		AppUtil.fillRow(row,styleValue,3,"");
		AppUtil.fillRow(row,styleValue,4,"");
		i=i+2;
		row = sheet.createRow((short) i);
		String color = Constants.COLOR3;
		AppUtil.fillRow(row,style,1,"Source Table");
		AppUtil.fillRow(row,style,2,"Source Column");
		AppUtil.fillRow(row,style,3,"Destination Table");
		AppUtil.fillRow(row,style,4,"Destination Column");
		i++;
		for(XmlSrcDestTableMap list : xmlData.getXmlSrcDestTableMap()){
			row = sheet.createRow((short) i);
			if(color.equals(Constants.COLOR3)) {
				style=MetadataStyleSheet.getCellStyleGray(hwb,0);
				color = Constants.COLOR2;
			}
			else {
				style=MetadataStyleSheet.getCellStyleBlue(hwb,0);
				color = Constants.COLOR3;
			}
			AppUtil.fillRow(row,style,1,list.getSrcTable());
			AppUtil.fillRow(row,style,2,list.getSrcColumn());
			AppUtil.fillRow(row,style,3,list.getDestTable());
			AppUtil.fillRow(row,style,4,list.getDestColumn());
			i++;
		}
	}
	
	

	public int dTDetailPopulation(HSSFWorkbook hwb,HSSFSheet sheet,StringBuilder dtData,int i) {
		HSSFCellStyle styleValue = MetadataStyleSheet.getCellStyleWhite(hwb,0);
		styleValue.setWrapText(true);
		String str = dtData.toString();
		int len =str.length();
		int start=0;
		int end=0;
		int range=400;
		//int i=5;
		int j;
		for(j = 1;j<=(len/range)+1;j++){
			HSSFRow row = sheet.createRow((short) i);
			sheet.autoSizeColumn(2);
			if((len-start)>range)
				end=range*j;
			else
				end=len;
			HSSFCell cell = row.createCell(2);
			cell.setCellValue(str.substring(start, end));
			cell.setCellStyle(styleValue);
			i++;
			start=range*j;
		}
		return j+i;
	}
}
