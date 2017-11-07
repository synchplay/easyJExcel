package org.easyj.easyjexcel;

import java.io.IOException;
import java.util.Collections;
import java.util.List;

import javax.servlet.http.HttpServletResponse;

import org.apache.poi.ss.usermodel.Workbook;
import org.easyj.easyjexcel.annotation.ExcelResponseBody;
import org.easyj.easyjexcel.constant.ExcelType;
import org.easyj.easyjexcel.converter.ExcelConverter;
import org.springframework.core.MethodParameter;
import org.springframework.web.context.request.NativeWebRequest;
import org.springframework.web.method.support.HandlerMethodReturnValueHandler;
import org.springframework.web.method.support.ModelAndViewContainer;

public class ExcelResponseBodyHandler implements HandlerMethodReturnValueHandler {

	private ExcelConverter converters;

	public void setConverters(ExcelConverter converters) {
		this.converters = converters;
	}

	@Override
	public boolean supportsReturnType(MethodParameter returnType) {
		return returnType.hasMethodAnnotation(ExcelResponseBody.class);
	}

	@Override
	public void handleReturnValue(Object returnValue, MethodParameter returnType, ModelAndViewContainer mavContainer,
			NativeWebRequest webRequest) throws Exception {
		HttpServletResponse servletResponse = webRequest.getNativeResponse(HttpServletResponse.class);

		ExcelResponseBody annotation = returnType.getMethodAnnotation(ExcelResponseBody.class);
		ExcelType type = annotation.type();
		String fileName = new String(annotation.name().getBytes(), "iso-8859-1") + type.suffix();

		servletResponse.setContentType("application/vnd.ms-converter");
		servletResponse.setHeader("content-disposition", "attachment;filename=" + fileName);
		List<?> excel;
		if (returnValue instanceof List) {
			excel = (List<?>) returnValue;
		} else {
			excel = Collections.singletonList(returnValue);
		}

		if (converters.supportsExcelType(type)) {
			try (Workbook workbook = converters.toExcel(annotation, excel)) {
				workbook.write(servletResponse.getOutputStream());
				servletResponse.flushBuffer();
			} catch (IOException ignored) {
				;
			}
		}

	}

}
