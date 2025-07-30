package com.main.excel.gui;

import java.io.FileOutputStream;
import java.sql.ResultSet;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import com.main.excel.app.Configurations;
import com.main.excel.app.Constants;
import com.main.excel.dao.MetadataDao;
import com.main.excel.model.T2TMapping;
import com.main.excel.service.CreateMetadataSheet;
import com.main.excel.service.MetadataDetailSheet;
import com.main.excel.service.MetadataStyleSheet;
import com.main.excel.util.AppUtil;
import com.trivadis.unwrapper.util.Unwrapper;

public class CreateRunChart {
	
	static String DESTINATION_DIR = Configurations.getInstance().getProperty(Constants.DESTINATION_DIR);
	//static String SOURCE_NAME = Configurations.getInstance().getProperty(Constants.SOURCE_NAME);
	
	public static void main(String[] args){
		CreateRunChart ct = new CreateRunChart();
		//ct.getExcel("1746101679792","R","8.1"); //CONTRACTUAL
		ct.getExcel("1753281862940","R","8.1"); //BAU
		
		//ct.getExcel("LOAD_FSI","B");
	}
	 
	public boolean getExcel(String runID,String type,String ofsaaVersion) {
		try {
			System.out.println("Process Start Time:"+AppUtil.showTime());
			MetadataDao metadata = new MetadataDao();
			ResultSet rs0 = metadata.getRunDetails(runID);
			String ExcelFileName = "\\";
			String RunName = null;
			List<String> functionList= new ArrayList<String>();
			while (rs0.next()) {
				RunName = rs0.getString("V_OBJECT_DESC");
			}
			if(RunName==null){
				RunName=runID;
			}
			//System.out.println("Date ::"+AppUtil.showDate());
			ExcelFileName = DESTINATION_DIR+ExcelFileName+RunName+"_"+AppUtil.showDate()+".xls";
			HSSFWorkbook hwb = MetadataStyleSheet.hwb;
			CreateMetadataSheet createMetadataSheet = new CreateMetadataSheet();
			MetadataDetailSheet metadataDetailSheet = new MetadataDetailSheet();
			//Biff8EncryptionKey.setCurrentUserPassword("5624");
			HSSFSheet sheet = createMetadataSheet.getCreateRunSheet(hwb,RunName);
			HSSFRow rowhead = sheet.createRow((short) 1);
			HSSFCellStyle style = new CreateMetadataSheet().getCreateStyle(hwb, rowhead);
			HSSFCellStyle style1 = new CreateMetadataSheet().getCreateStyle(hwb, rowhead);
			int i = 2;
			String str1 = new String("ABC");
			String str2 = new String(Constants.COLOR3);
			HSSFCell cell = null;
			HSSFRow row = null;
			ResultSet rs1 = metadata.getTaskDetails(runID,type);
			int totalTask=AppUtil.getRowSetSize(rs1);
			int percent=0;
			functionList = metadata.getAllFunctions();
			while (rs1.next()){
				if (!(str1.equals(rs1.getString("PROCESS_NAME")) || i == 2 )) {
					if (str2.equals(Constants.COLOR2)) {
						style = MetadataStyleSheet.getCellStyleBlue(hwb, 0);
					} else {
						style = MetadataStyleSheet.getCellStyleGray(hwb, 0);
					}
					row = sheet.createRow((short) i);
					row.createCell(1).setCellStyle(style);
					row.createCell(2).setCellStyle(style);
					row.createCell(3).setCellStyle(style);
					row.createCell(4).setCellStyle(style);
					row.createCell(5).setCellStyle(style);
					row.createCell(6).setCellStyle(style);
					row.createCell(7).setCellStyle(style);
					i++;
				}
				row = sheet.createRow((short) i);

				if ((i % 2) == 1) {
					style = MetadataStyleSheet.getCellStyleGray(hwb, 0);
					style1 = MetadataStyleSheet.getCellStyleGray_UnderLine(hwb);
					str2 = Constants.COLOR2;
				} else {
					style = MetadataStyleSheet.getCellStyleBlue(hwb, 0);
					style1 = MetadataStyleSheet.getCellStyleBlue_UnderLine(hwb);
					str2 = Constants.COLOR3;
				}
				
				cell = row.createCell(2);
				if (!str1.equals(rs1.getString("PROCESS_NAME"))) {
					cell.setCellValue(rs1.getString("PROCESS_NAME"));
				}
				cell.setCellStyle(style);
				
				HSSFSheet sheet1;
				ResultSet rs3;
				String taskType= rs1.getString("TASK_TYPE");
				String taskName= rs1.getString("TASK_NAME");
				String taskID= rs1.getString("TASK_ID");
				String sheetName = AppUtil.getSheetName(taskName)+i;
				createMetadataSheet.createACell(hwb, row, rs1.getString("TASK_SEQ"), style, 1, sheetName);
				
				/* Rule*/
				if (taskType.startsWith("TYPE")) {
					//System.out.println("Task Name:"+taskName+" : "+AppUtil.showTime());
					int rowid = 0;
					sheet1 = createMetadataSheet.getCreateRuleSheet(hwb, sheetName);
					rs3 = metadata.getRuleDetails(taskID);
					rowid = metadataDetailSheet.ruleDetailPopulation(hwb, sheet1, rs3);
					rs3 = metadata.getRuleSourceTargetDetails(taskID);
					rowid = metadataDetailSheet.ruleHireBPPopulation(hwb, sheet1, rs3, rowid);
					List<String> arrayList = AppUtil.getSrcTargetHier(metadata.getRuleSourceTargetDetails(taskID));
					rs3 = metadata.getRuleMappingDetails(taskID);
					arrayList = metadataDetailSheet.ruleMappingPopulation(hwb, sheet1, rs3, rowid,arrayList);
					rowid = Integer.parseInt(arrayList.get(arrayList.size() - 1)) + 3;
					if (arrayList.size() > 1) {
						for (String temp : arrayList) {
							rs3 = metadata.getBPDetails(temp);
							rowid = metadataDetailSheet.ruleBPDetailPopulation(hwb, sheet1, rs3, rowid);
							rowid = rowid + 2;
						}
					}
					rs3.close();
				}
				/*Data Transformation*/
				else if(taskType.startsWith("DATE") && !taskName.contains("Agg_Cash_Flows_Populate")) {
					//System.out.print("Task Name:"+taskName+" : "+AppUtil.showTime());
					sheet1 = createMetadataSheet.getCreateDTSheet(hwb, sheetName,taskName);
					rs3 = metadata.getDTDetails(taskName);
					StringBuilder wrapText= new StringBuilder();
					while (rs3.next()) {
						wrapText.append(rs3.getString("TEXT"));
					}
					StringBuilder unWrapText = new StringBuilder();
					if (wrapText.indexOf("wrapped")!=-1){
						unWrapText.append(Unwrapper.unwrap(wrapText.toString()));
					}
					else
						unWrapText.append(wrapText);
					int rowid = metadataDetailSheet.dTDetailPopulation(hwb, sheet1, unWrapText,5);
					StringBuilder unWrapTextObj =AppUtil.getObjectName(unWrapText,functionList);
					metadataDetailSheet.dTDetailPopulation(hwb, sheet1, unWrapTextObj,rowid+5);
					rs3.close();
					//System.out.println("Task ::"+taskName+" -->Process End Time:"+AppUtil.showTime());
				}
				/*T2T*/
				else if(taskType.startsWith("T2T")) {
					//System.out.println("Task Name:"+taskName+" : "+AppUtil.showTime());
					sheet1 = createMetadataSheet.getCreateT2TSheet(hwb, sheetName,taskName);
					T2TMapping t2tMapping = new T2TMapping();
					if(ofsaaVersion.equals("8.0")) {
						t2tMapping=AppUtil.getT2TDetailsXml(hwb,taskName,sheetName);
					}
					else {
						t2tMapping=AppUtil.getT2TDetails(hwb,taskName,sheetName);
					}
					//System.out.println("Task Name :"+taskName);
					metadataDetailSheet.t2TDetailPopulation(hwb, sheet1, t2tMapping);
				}
				//System.out.println(" End : "+AppUtil.showTime());
				//createMetadataSheet.createACell(hwb, row, rs1.getString("PROCESS_NAME"), style, 2, sheetName);
				createMetadataSheet.createACell(hwb, row, rs1.getString("TASK_NAME"), style1, 3, sheetName);
				createMetadataSheet.createACell(hwb, row, rs1.getString("TASK_TYPE"), style, 4, sheetName);
				createMetadataSheet.createACell(hwb, row, rs1.getString("TASK_PRECEDENCE"), style, 5, sheetName);
				createMetadataSheet.createACell(hwb, row, null, style, 6, sheetName);
				createMetadataSheet.createACell(hwb, row, null, style, 7, sheetName);
				str1 = rs1.getString("PROCESS_NAME");
				i++;
				if(((i*100)/totalTask)>=(percent+1) && ((i*100)/totalTask) <90){
					percent = (i*100)/totalTask+9;
					System.out.print(percent+"%  ");
				}
				else if(i==totalTask)
					System.out.print("100%  ");
				
			}
			FileOutputStream fileOut = new FileOutputStream(ExcelFileName);
			//hwb.writeProtectWorkbook(Biff8EncryptionKey.getCurrentUserPassword(), "");
			hwb.write(fileOut);
			fileOut.close();
			System.out.println("\nYour excel file has been generated!");
			rs1.close();
			
			System.out.println("Process End Time:"+AppUtil.showTime());
		} catch (Exception ex) {
			ex.printStackTrace();
			System.out.println(ex);
			return false;
		}
		return true;
	}
	
}