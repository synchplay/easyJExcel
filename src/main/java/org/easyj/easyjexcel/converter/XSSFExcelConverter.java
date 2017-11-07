package org.easyj.easyjexcel.converter;

import java.io.InputStream;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.easyj.easyjexcel.constant.ExcelType;

public class XSSFExcelConverter extends AbstractGenericPoiExcelConverter {

    @Override
    public boolean supportsExcelType(ExcelType excelType) {
        return ExcelType.XLSX == excelType;
    }

    @Override
    protected Workbook createWorkBook(InputStream inputStream) throws Exception {
        return new XSSFWorkbook(inputStream);
    }

    @Override
    protected Workbook createWorkBook() throws Exception {
        return new XSSFWorkbook();
    }
}