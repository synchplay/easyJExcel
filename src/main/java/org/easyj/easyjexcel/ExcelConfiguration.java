package org.easyj.easyjexcel;

import java.util.List;

import org.easyj.easyjexcel.converter.ExcelConverter;
import org.easyj.easyjexcel.converter.ExcelConverterComposite;
import org.springframework.context.annotation.Configuration;
import org.springframework.web.method.support.HandlerMethodArgumentResolver;
import org.springframework.web.method.support.HandlerMethodReturnValueHandler;
import org.springframework.web.servlet.config.annotation.WebMvcConfigurerAdapter;

@Configuration
public class ExcelConfiguration extends WebMvcConfigurerAdapter {

	private final ExcelConverter excelConverter = new ExcelConverterComposite();

	@Override
	public void addArgumentResolvers(List<HandlerMethodArgumentResolver> argumentResolvers) {
		ExcelRequestBodyHandler defaultExcelHandler = new ExcelRequestBodyHandler();
		defaultExcelHandler.setConverters(excelConverter);
		argumentResolvers.add(defaultExcelHandler);
	}

	@Override
	public void addReturnValueHandlers(List<HandlerMethodReturnValueHandler> returnValueHandlers) {
		ExcelResponseBodyHandler defaultExcelHandler = new ExcelResponseBodyHandler();
		defaultExcelHandler.setConverters(excelConverter);
		returnValueHandlers.add(defaultExcelHandler);
	}

}
