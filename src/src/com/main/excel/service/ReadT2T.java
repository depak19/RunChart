package com.main.excel.service;

import java.sql.ResultSet;
import java.util.ArrayList;
import java.util.List;

import org.apache.commons.lang3.StringUtils;

import com.main.excel.dao.MetadataDao;
import com.main.excel.model.T2TMapping;
import com.main.excel.model.XmlSrcDestTableMap;

public class ReadT2T {
	
	public ReadT2T() {
		super();
	}
	public static T2TMapping getT2TMapping(String t2tName) {
		T2TMapping xmlData = new T2TMapping();
		MetadataDao metadata = new MetadataDao();
		List<XmlSrcDestTableMap> xmlSrcDestTableMapList = new ArrayList<XmlSrcDestTableMap>();
		ResultSet rs1;
		try {
			rs1=metadata.getT2TDetails(t2tName);
			boolean FLAG = true;
			while (rs1.next()) {
			XmlSrcDestTableMap xmlSrcDestTableMap = new XmlSrcDestTableMap();
			xmlSrcDestTableMap.setSrcTable(rs1.getString("V_SRC_TABLE_ID"));
			if(rs1.getString("V_SRC_TABLE_ID").equals("EXPRESSION"))
				xmlSrcDestTableMap.setSrcColumn(rs1.getString("V_EXPRESSION"));
			else
				xmlSrcDestTableMap.setSrcColumn(rs1.getString("V_SRC_COLUMN_ID"));
			xmlSrcDestTableMap.setSrcExpression(rs1.getString("V_EXPRESSION"));
			xmlSrcDestTableMap.setDestTable(rs1.getString("V_DEST_TABLE_ID"));
			xmlSrcDestTableMap.setDestColumn(rs1.getString("V_DEST_COLUMN_ID"));
			xmlSrcDestTableMapList.add(xmlSrcDestTableMap);
			if(FLAG) {
				String ansiJoin = StringUtils.normalizeSpace(rs1.getString("V_ANSIJOIN")).replace("\n"," ").toUpperCase();
				xmlData.setAnsiJoin(ansiJoin);
				xmlData.setFilter(rs1.getString("V_FILTER") != null ? StringUtils.normalizeSpace(rs1.getString("V_FILTER")).replace("\n"," ").toUpperCase() : "");
				FLAG=false;
				}
			}
			xmlData.setXmlSrcDestTableMap(xmlSrcDestTableMapList);
			rs1.close();
		}
		catch (Exception e) {
			System.out.println("T2T Not Found::"+t2tName);
		}
		return xmlData;
	}
}
