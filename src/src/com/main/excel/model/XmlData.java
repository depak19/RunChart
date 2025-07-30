package com.main.excel.model;

import java.io.Serializable;
import java.util.ArrayList;
import java.util.List;

public class XmlData implements Serializable{

	/**
	 * 
	 */
	private static final long serialVersionUID = 1L;
	private String filter;
	private String ansiJoin;
	private List <XmlSrcDestTableMap> xmlSrcDestTableMap = new ArrayList<XmlSrcDestTableMap>();
	
	public XmlData () {
		super();
	}

	public String getFilter() {
		return filter;
	}

	public void setFilter(String filter) {
		this.filter = filter;
	}

	public String getAnsiJoin() {
		return ansiJoin;
	}

	public void setAnsiJoin(String ansiJoin) {
		this.ansiJoin = ansiJoin;
	}

	public List<XmlSrcDestTableMap> getXmlSrcDestTableMap() {
		return xmlSrcDestTableMap;
	}

	public void setXmlSrcDestTableMap(List<XmlSrcDestTableMap> xmlSrcDestTableMap) {
		this.xmlSrcDestTableMap = xmlSrcDestTableMap;
	}
	
}
