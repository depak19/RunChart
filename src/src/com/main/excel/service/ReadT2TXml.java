package com.main.excel.service;

import java.io.File;
import java.io.FileNotFoundException;
import java.util.ArrayList;
import java.util.List;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;

import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;

import com.main.excel.app.Configurations;
import com.main.excel.app.Constants;
import com.main.excel.model.T2TMapping;
import com.main.excel.model.XmlSrcDestTableMap;

public class ReadT2TXml {
	static String joinDir = "\\ftpshare\\STAGE\\";
	static String mappingDir = "\\ftpshare\\"+Configurations.getInstance().getProperty(Constants.INFODOM)+"\\erwin\\Sources\\";
	
	public ReadT2TXml() {
		super();
	}
	public static T2TMapping readXMLFileCreateXLS(String dir,String XmlFile,String source) throws FileNotFoundException {
		String fileName =dir+mappingDir+source+"\\"+XmlFile+".mapping.xml";
		File fXmlFile = new File(fileName);
		T2TMapping xmlData = new T2TMapping();
		try {
			DocumentBuilderFactory dbFactory = DocumentBuilderFactory.newInstance();
			DocumentBuilder dBuilder = dbFactory.newDocumentBuilder();
			Document doc = dBuilder.parse(fXmlFile);
			doc.getDocumentElement().normalize();
			String descTableName = null;

			NodeList nList1 = doc.getElementsByTagName("TABLE");
			for (int i = 0; i < nList1.getLength(); i++) {
				Node nNodeD = nList1.item(i);
				descTableName = ((Element) nNodeD).getAttribute("ID");
			}
			NodeList nList = doc.getElementsByTagName("MAP");
			List<XmlSrcDestTableMap> xmlSrcDestTableMapList = new ArrayList<XmlSrcDestTableMap>();
			for (int i = 0; i < nList.getLength(); i++) {
				Node nNode = nList.item(i);
				if (nNode.getNodeType() == Node.ELEMENT_NODE) {
					Element eElement = (Element) nNode;
					NodeList nodeSrc = eElement.getElementsByTagName("SOURCE").item(0).getChildNodes();
					NodeList nodesDest = eElement.getElementsByTagName("DESTINATION").item(0).getChildNodes();
					for (int j = 0; j < nodeSrc.getLength(); j++) {
						
						Element eElementSrc = (Element) nodeSrc;
						boolean FLAG = (eElementSrc.getAttribute("TABLENAME").equals("EXPRESSION") ? true : false);
						Node nodeS = (Node) nodeSrc.item(j);
						Node nodeD = (Node) nodesDest.item(j);
						if (nodeS.getNodeType() == Node.ELEMENT_NODE) {
							XmlSrcDestTableMap xmlSrcDestTableMap = new XmlSrcDestTableMap();
							Element eElementS = (Element) nodeS;
							Element eElementD = (Element) nodeD;
							xmlSrcDestTableMap.setDestTable(descTableName);
							xmlSrcDestTableMap.setDestColumn(eElementD.getAttribute("ID"));
							xmlSrcDestTableMap.setSrcTable(eElementSrc.getAttribute("TABLENAME"));
							if (FLAG) {
								xmlSrcDestTableMap.setSrcColumn(eElementS.getElementsByTagName("EXPRESSION").item(0).getTextContent());
							}
							else {
								xmlSrcDestTableMap.setSrcColumn(eElementS.getAttribute("ID"));
							}
							
							xmlSrcDestTableMapList.add(xmlSrcDestTableMap);
						}
						
					}
					xmlData.setXmlSrcDestTableMap(xmlSrcDestTableMapList);
				}

			}
			fileName =dir+joinDir+source+"\\"+XmlFile+".xml";
			fXmlFile = new File(fileName);
			dbFactory = DocumentBuilderFactory.newInstance();
			dBuilder = dbFactory.newDocumentBuilder();
			doc = dBuilder.parse(fXmlFile);
			doc.getDocumentElement().normalize();
			NodeList nList2 = doc.getElementsByTagName("DEFINITION");
			for (int i = 0; i < nList2.getLength(); i++) {
				Node nNodeFJ = nList2.item(i);
				if (nNodeFJ.getNodeType() == Node.ELEMENT_NODE) {
					Element eElementFJ = (Element) nNodeFJ;
						xmlData.setFilter(eElementFJ.getElementsByTagName("FILTER").item(0).getTextContent().toUpperCase());
						xmlData.setAnsiJoin(eElementFJ.getElementsByTagName("ANSIJOIN").item(0).getTextContent().toUpperCase());
					}
			}
		} catch (Exception e) {
			System.out.println("File Not Found::"+fileName);
		}
		return xmlData;
	}

}