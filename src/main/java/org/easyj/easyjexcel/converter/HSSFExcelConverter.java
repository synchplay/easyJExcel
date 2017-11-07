package org.easyj.easyjexcel.converter;

import java.io.InputStream;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.easyj.easyjexcel.constant.ExcelType;

public class HSSFExcelConverter extends AbstractGenericPoiExcelConverter {

    @Override
    public boolean supportsExcelType(ExcelType excelType) {
        return ExcelType.XLS == excelType;
    }

    @Override
    protected Workbook createWorkBook(InputStream inputStream) throws Exception{
        return new HSSFWorkbook(inputStream);
    }

    @Override
    protected Workbook createWorkBook() throws Exception {
        return new HSSFWorkbook();
    }
}