package com.main.excel.model;

import java.io.Serializable;

public class XmlSrcDestTableMap implements Serializable {
	/**
	 * 
	 */
	private static final long serialVersionUID = 1L;
	private String srcTable;
	private String srcColumn;
	private String destTable;
	private String destColumn;
	private String srcExpression;
	
	public XmlSrcDestTableMap() {
		super();
	}

	public String getSrcTable() {
		return srcTable;
	}

	public void setSrcTable(String srcTable) {
		this.srcTable = srcTable;
	}

	public String getSrcColumn() {
		return srcColumn;
	}

	public void setSrcColumn(String srcColumn) {
		this.srcColumn = srcColumn;
	}

	public String getDestTable() {
		return destTable;
	}

	public void setDestTable(String destTable) {
		this.destTable = destTable;
	}

	public String getDestColumn() {
		return destColumn;
	}

	public void setDestColumn(String destColumn) {
		this.destColumn = destColumn;
	}

	public String getSrcExpression() {
		return srcExpression;
	}

	public void setSrcExpression(String srcExpression) {
		this.srcExpression = srcExpression;
	}

}
