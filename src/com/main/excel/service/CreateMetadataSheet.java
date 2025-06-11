package com.main.excel.service;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFHyperlink;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.util.CellRangeAddress;

import com.main.excel.app.Constants;

public class CreateMetadataSheet {
	
	public HSSFSheet getCreateRunSheet(HSSFWorkbook hwb, String RunName) {
		HSSFSheet sheet = hwb.createSheet(Constants.RUN_CHART);
		sheet.setColumnWidth(1, (8*256));
		sheet.setColumnWidth(2, (40*256));
		sheet.setColumnWidth(3, (40*256));
		sheet.setColumnWidth(4, (10*256));
		sheet.setColumnWidth(5, (40*256));
		sheet.setColumnWidth(6, (10*256));
		sheet.setColumnWidth(7, (50*256));
		HSSFRow row = sheet.createRow((short) 0);
		HSSFCellStyle style = MetadataStyleSheet.getCellStyleHeaderH1(hwb);
		HSSFCell cell = row.createCell(1);
		cell.setCellStyle(style);
		String mergeCell ="B"+1+":H"+1;
		sheet.addMergedRegion(CellRangeAddress.valueOf(mergeCell));
		cell.setCellValue(RunName);
		cell.setCellStyle(style);
		return sheet;
	}
	public void createACell (HSSFWorkbook hwb,HSSFRow row,String value,HSSFCellStyle style,int i,String link) {
		HSSFCell cell = null;
		HSSFHyperlink url_link=new HSSFHyperlink(HSSFHyperlink.LINK_DOCUMENT);
		String str = link+"!A1";
		cell = row.createCell(i);
		if (i==3) {
			url_link.setAddress(str);
			cell.setHyperlink(url_link);
		}
		cell.setCellValue(value);
		cell.setCellStyle(style);
	}
	public HSSFSheet getCreateRuleSheet(HSSFWorkbook hwb,String ruleSheet) {
		HSSFHyperlink url_link=new HSSFHyperlink(HSSFHyperlink.LINK_DOCUMENT);
		String str = "'"+Constants.RUN_CHART+"'!A1";
		HSSFSheet sheet = hwb.createSheet(ruleSheet);
		HSSFRow row = sheet.createRow((short) 0);
		HSSFCellStyle style = MetadataStyleSheet.getCellBackToIndex(hwb);
		HSSFCell cell = row.createCell(0);
		sheet.setColumnWidth(0, (20*256));
		sheet.setColumnWidth(1, (60*256));
		sheet.setColumnWidth(2, (60*256));
		cell.setCellStyle(style);
		url_link.setAddress(str);
		cell.setHyperlink(url_link);
		cell.setCellValue("Go back to Index");
		return sheet;
	}
	
	public HSSFSheet getCreateT2TSheet(HSSFWorkbook hwb,String t2TSheet,String t2TName) {
		HSSFHyperlink url_link=new HSSFHyperlink(HSSFHyperlink.LINK_DOCUMENT);
		String str = "'"+Constants.RUN_CHART+"'!A1";
		HSSFSheet sheet = hwb.createSheet(t2TSheet);
		HSSFRow row = sheet.createRow((short) 0);
		HSSFCellStyle style = MetadataStyleSheet.getCellBackToIndex(hwb);
		HSSFCell cell = row.createCell(0);
		sheet.setColumnWidth(0, (25*256));
		sheet.setColumnWidth(1, (30*256));
		sheet.setColumnWidth(2, (50*256));
		sheet.setColumnWidth(3, (30*256));
		sheet.setColumnWidth(4, (30*256));
		cell.setCellStyle(style);
		url_link.setAddress(str);
		cell.setHyperlink(url_link);
		cell.setCellValue("Go back to Index");
		style =MetadataStyleSheet.getCellStyleHeaderH1(hwb);
		row = sheet.createRow((short) 4);
		cell = row.createCell(1);
		String mergeCell ="B"+5+":E"+5;
		sheet.addMergedRegion(CellRangeAddress.valueOf(mergeCell));
		cell.setCellValue(t2TName);
		cell.setCellStyle(style);
		cell = row.createCell(2);
		cell.setCellStyle(style);
		cell = row.createCell(3);
		cell.setCellStyle(style);
		cell = row.createCell(4);
		cell.setCellStyle(style);
		return sheet;
	}
	
	public HSSFSheet getCreateDTSheet(HSSFWorkbook hwb,String t2TSheet,String t2TName) {
		HSSFHyperlink url_link=new HSSFHyperlink(HSSFHyperlink.LINK_DOCUMENT);
		String str = "'"+Constants.RUN_CHART+"'!A1";
		HSSFSheet sheet = hwb.createSheet(t2TSheet);
		HSSFRow row = sheet.createRow((short) 0);
		HSSFCellStyle style = MetadataStyleSheet.getCellBackToIndex(hwb);
		HSSFCell cell = row.createCell(0);
		cell.setCellStyle(style);
		url_link.setAddress(str);
		cell.setHyperlink(url_link);
		cell.setCellValue("Go back to Index");
		return sheet;
	}
	public HSSFCellStyle getCreateStyle(HSSFWorkbook hwb,HSSFRow rowhead) {
		HSSFCellStyle style= MetadataStyleSheet.getCellStyleHeaderH2(hwb);
		
		HSSFCell cell1 = rowhead.createCell(1);
		cell1.setCellValue("Task Seq");
		cell1.setCellStyle(style);
		
		HSSFCell cell2 = rowhead.createCell(2);
		cell2.setCellValue("Process Tree");
		cell2.setCellStyle(style);
		
		HSSFCell cell3 = rowhead.createCell(3);
		cell3.setCellValue("Rule");
		cell3.setCellStyle(style);
		
		HSSFCell cell4 = rowhead.createCell(4);
		cell4.setCellValue("Component");
		cell4.setCellStyle(style);

		HSSFCell cell5 = rowhead.createCell(5);
		cell5.setCellValue("Precedence");
		cell5.setCellStyle(style);
		
		HSSFCell cell6 = rowhead.createCell(6);
		cell6.setCellValue("PID");
		cell6.setCellStyle(style);
		
		HSSFCell cell7 = rowhead.createCell(7);
		cell7.setCellValue("Comments");
		cell7.setCellStyle(style);
		return style;
	}

}
