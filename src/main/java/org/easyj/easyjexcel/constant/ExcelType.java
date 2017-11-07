package org.easyj.easyjexcel.constant;

public enum ExcelType {

	XLS(".xls"),

	XLSX(".xlsx");

	private String suffix;

	ExcelType(String suffix) {
		this.suffix = suffix;
	}

	public String suffix() {
		return suffix;
	}
}