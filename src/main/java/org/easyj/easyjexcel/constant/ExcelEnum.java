package org.easyj.easyjexcel.constant;

public interface ExcelEnum<T> {

	/**
	 * Excel显示名称
	 * 
	 * @return
	 */
	String excelName();

	/**
	 * 值
	 * 
	 * @return
	 */
	T excelValue();

}
