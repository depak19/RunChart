package com.main.excel.service;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;

public class MetadataStyleSheet {
	static public HSSFWorkbook hwb = new HSSFWorkbook();
	static public HSSFCellStyle styleBackIndex = hwb.createCellStyle();
	static public HSSFFont fontBackIndex = hwb.createFont();
	static public HSSFCellStyle styleH2 = hwb.createCellStyle();
	static public HSSFFont fontH2 = hwb.createFont();
	static public HSSFCellStyle styleBlue = hwb.createCellStyle();
	static public HSSFFont fontBlue = hwb.createFont();
	static public HSSFCellStyle styleGray = hwb.createCellStyle();
	static public HSSFFont fontGray = hwb.createFont();
	static public HSSFCellStyle styleH1 = hwb.createCellStyle();
	static public HSSFFont fontH1 = hwb.createFont();
	static public HSSFCellStyle styleBlueUnderLine = hwb.createCellStyle();
	static public HSSFFont fontBlueUnderline = hwb.createFont();
	static public HSSFCellStyle styleGrayUnderLine = hwb.createCellStyle();
	static public HSSFFont fontGrayUnderline = hwb.createFont();
	static public HSSFCellStyle styleWhite = hwb.createCellStyle();
	static public HSSFFont fontWhite = hwb.createFont();
	static public String FONT_CALIBRI = "CALIBRI";
	static public String FONT_TAHOMA = "TAHOMA";
	static public int SIZE_8 = 8;
	static public int SIZE_10 = 10;
	//static public int SIZE_12 = 12;
	//static public int SIZE_14 = 14;

	public MetadataStyleSheet() {
		super();
	}

	public static HSSFCellStyle getCellBackToIndex(HSSFWorkbook hwb) {
		styleBackIndex.setFillForegroundColor(HSSFColor.WHITE.index);
		styleBackIndex.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
		styleBackIndex.setBorderTop(HSSFCellStyle.BORDER_THIN);
		styleBackIndex.setBorderBottom(HSSFCellStyle.BORDER_THIN);
		styleBackIndex.setBorderLeft(HSSFCellStyle.BORDER_THIN);
		styleBackIndex.setBorderRight(HSSFCellStyle.BORDER_THIN);
		fontBackIndex.setColor(IndexedColors.BLUE.getIndex());
		fontBackIndex.setUnderline(Font.U_SINGLE);
		fontBackIndex.setFontName(FONT_CALIBRI);
		fontBackIndex.setFontHeightInPoints((short) SIZE_10);
		styleBackIndex.setFont(fontBackIndex);
		return styleBackIndex;
	}

	public static HSSFCellStyle getCellStyleHeaderH2(HSSFWorkbook hwb) {
		styleH2.setFillForegroundColor(HSSFColor.GREEN.index);
		styleH2.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
		styleH2.setBorderTop(HSSFCellStyle.BORDER_THIN);
		styleH2.setBorderBottom(HSSFCellStyle.BORDER_THIN);
		styleH2.setBorderLeft(HSSFCellStyle.BORDER_THIN);
		styleH2.setBorderRight(HSSFCellStyle.BORDER_THIN);
		fontH2.setColor(HSSFColor.BLACK.index);
		fontH2.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
		fontH2.setFontName(FONT_CALIBRI);
		fontH2.setColor(HSSFColor.WHITE.index);
		fontH2.setFontHeightInPoints((short) SIZE_10);
		styleH2.setFont(fontH2);
		return styleH2;
	}

	public static HSSFCellStyle getCellStyleGray(HSSFWorkbook hwb, int i) {
		styleGray.setFillForegroundColor(HSSFColor.GREY_25_PERCENT.index);
		styleGray.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
		styleGray.setBorderTop(HSSFCellStyle.BORDER_THIN);
		styleGray.setBorderBottom(HSSFCellStyle.BORDER_THIN);
		styleGray.setBorderLeft(HSSFCellStyle.BORDER_THIN);
		styleGray.setBorderRight(HSSFCellStyle.BORDER_THIN);
		//styleGray.setWrapText(true);
		fontGray.setColor(HSSFColor.BLACK.index);
		fontGray.setFontName(FONT_CALIBRI);
		fontGray.setFontHeightInPoints((short) SIZE_8);
		styleGray.setFont(fontGray);
		return styleGray;
	}

	public static HSSFCellStyle getCellStyleBlue(HSSFWorkbook hwb, int i) {
		styleBlue.setFillForegroundColor(HSSFColor.PALE_BLUE.index);
		//styleBlue.setFillForegroundColor(new XSSFColor(new java.awt.Color(128, 0, 128)));
		styleBlue.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
		styleBlue.setBorderTop(HSSFCellStyle.BORDER_THIN);
		styleBlue.setBorderBottom(HSSFCellStyle.BORDER_THIN);
		styleBlue.setBorderLeft(HSSFCellStyle.BORDER_THIN);
		styleBlue.setBorderRight(HSSFCellStyle.BORDER_THIN);
		//styleBlue.setWrapText(true);
		fontBlue.setColor(HSSFColor.BLACK.index);
		fontBlue.setFontName(FONT_CALIBRI);
		fontBlue.setFontHeightInPoints((short) SIZE_8);
		styleBlue.setFont(fontBlue);
		return styleBlue;
	}

	public static HSSFCellStyle getCellStyleGray_UnderLine(HSSFWorkbook hwb) {
		styleGrayUnderLine.setFillForegroundColor(HSSFColor.GREY_25_PERCENT.index);
		styleGrayUnderLine.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
		styleGrayUnderLine.setBorderTop(HSSFCellStyle.BORDER_THIN);
		styleGrayUnderLine.setBorderBottom(HSSFCellStyle.BORDER_THIN);
		styleGrayUnderLine.setBorderLeft(HSSFCellStyle.BORDER_THIN);
		styleGrayUnderLine.setBorderRight(HSSFCellStyle.BORDER_THIN);
		//styleGrayUnderLine.setWrapText(true);
		fontGrayUnderline.setFontName(FONT_CALIBRI);
		fontGrayUnderline.setFontHeightInPoints((short) SIZE_8);
		fontGrayUnderline.setUnderline(Font.U_SINGLE);
		fontGrayUnderline.setColor(IndexedColors.BLUE.getIndex());
		styleGrayUnderLine.setFont(fontGrayUnderline);
		return styleGrayUnderLine;
	}

	public static HSSFCellStyle getCellStyleBlue_UnderLine(HSSFWorkbook hwb) {
		styleBlueUnderLine.setFillForegroundColor(HSSFColor.PALE_BLUE.index);
		styleBlueUnderLine.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
		styleBlueUnderLine.setBorderTop(HSSFCellStyle.BORDER_THIN);
		styleBlueUnderLine.setBorderBottom(HSSFCellStyle.BORDER_THIN);
		styleBlueUnderLine.setBorderLeft(HSSFCellStyle.BORDER_THIN);
		styleBlueUnderLine.setBorderRight(HSSFCellStyle.BORDER_THIN);
		//styleBlueUnderLine.setWrapText(true);
		fontBlueUnderline.setFontName(FONT_CALIBRI);
		fontBlueUnderline.setFontHeightInPoints((short) SIZE_8);
		fontBlueUnderline.setUnderline(Font.U_SINGLE);
		fontBlueUnderline.setColor(IndexedColors.BLUE.getIndex());
		styleBlueUnderLine.setFont(fontBlueUnderline);
		return styleBlueUnderLine;
	}

	public static HSSFCellStyle getCellStyleHeaderH1(HSSFWorkbook hwb) {
		styleH1.setFillForegroundColor(HSSFColor.RED.index);
		styleH1.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
		styleH1.setBorderTop(HSSFCellStyle.BORDER_THIN);
		styleH1.setBorderBottom(HSSFCellStyle.BORDER_THIN);
		styleH1.setBorderLeft(HSSFCellStyle.BORDER_THIN);
		styleH1.setBorderRight(HSSFCellStyle.BORDER_THIN);
		styleH1.setAlignment(HSSFCellStyle.ALIGN_CENTER);
		fontH1.setColor(HSSFColor.BLACK.index);
		fontH1.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
		fontH1.setFontName(FONT_TAHOMA);
		fontH1.setColor(HSSFColor.WHITE.index);
		fontH1.setFontHeightInPoints((short) SIZE_10);
		styleH1.setFont(fontH1);
		return styleH1;
	}

	public static HSSFCellStyle getCellStyleWhite(HSSFWorkbook hwb, int i) {
		styleWhite.setFillForegroundColor(HSSFColor.PALE_BLUE.index);
		styleWhite.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
		styleWhite.setBorderTop(HSSFCellStyle.BORDER_THIN);
		styleWhite.setBorderBottom(HSSFCellStyle.BORDER_THIN);
		styleWhite.setBorderLeft(HSSFCellStyle.BORDER_THIN);
		styleWhite.setBorderRight(HSSFCellStyle.BORDER_THIN);
		fontWhite.setColor(HSSFColor.BLACK.index);
		fontWhite.setFontName(FONT_CALIBRI);
		fontWhite.setFontHeightInPoints((short) SIZE_10);
		styleWhite.setFont(fontWhite);
		return styleWhite;
	}
}
