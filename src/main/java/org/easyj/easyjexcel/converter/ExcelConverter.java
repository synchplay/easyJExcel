package org.easyj.easyjexcel.converter;

import java.io.InputStream;
import java.util.List;

import org.apache.poi.ss.usermodel.Workbook;
import org.easyj.easyjexcel.annotation.ExcelRequestBody;
import org.easyj.easyjexcel.annotation.ExcelResponseBody;
import org.easyj.easyjexcel.constant.ExcelType;

public interface ExcelConverter {

	boolean supportsExcelType(ExcelType excelType);

	List<?> fromExcel(ExcelRequestBody excelRequestBody, InputStream input) throws Exception;

	<T> Workbook toExcel(ExcelResponseBody excelResponseBody, List<T> excelVoList) throws Exception;

}