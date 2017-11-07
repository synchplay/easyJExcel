package org.easyj.easyjexcel.converter;

import java.io.InputStream;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.Workbook;
import org.easyj.easyjexcel.ExcelRequestBodyHandler;
import org.easyj.easyjexcel.annotation.ExcelRequestBody;
import org.easyj.easyjexcel.annotation.ExcelResponseBody;
import org.easyj.easyjexcel.constant.ExcelType;
import org.springframework.util.ClassUtils;

public class ExcelConverterComposite implements ExcelConverter {

	private static boolean HSSFPresent = ClassUtils.isPresent("org.apache.poi.hssf.usermodel.HSSFWorkbook",
			ExcelRequestBodyHandler.class.getClassLoader());

	private static boolean XSSFPresent = ClassUtils.isPresent("org.apache.poi.xssf.usermodel.XSSFWorkbook",
			ExcelRequestBodyHandler.class.getClassLoader());

	private final Map<ExcelType, ExcelConverter> excelConvertersCache = new HashMap<>(4);

	public ExcelConverterComposite() {
		if (HSSFPresent) {
			excelConvertersCache.put(ExcelType.XLS, new HSSFExcelConverter());
		}
		if (XSSFPresent) {
			excelConvertersCache.put(ExcelType.XLSX, new XSSFExcelConverter());
		}
	}

	@Override
	public boolean supportsExcelType(ExcelType excelType) {
		return (excelConvertersCache.get(excelType) != null);
	}

	@Override
	public List<?> fromExcel(ExcelRequestBody excelRequestBody, InputStream input) throws Exception {
		ExcelType excelType = excelRequestBody.type();
		ExcelConverter converter = excelConvertersCache.get(excelType);
		if (converter == null) {
			throw new IllegalArgumentException("Unknown converter excelType [" + excelType.name() + "]");
		}
		return converter.fromExcel(excelRequestBody, input);
	}

	@Override
	public <T> Workbook toExcel(ExcelResponseBody excelResponseBody, List<T> excelVoList) throws Exception {
		ExcelType excelType = excelResponseBody.type();
		ExcelConverter converter = excelConvertersCache.get(excelType);
		if (converter == null) {
			throw new IllegalArgumentException("Unknown converter type [" + excelType.name() + "]");
		}
		return converter.toExcel(excelResponseBody, excelVoList);
	}

}