package org.easyj.easyjexcel;

import java.util.ArrayList;
import java.util.List;

import javax.servlet.http.HttpServletRequest;

import org.easyj.easyjexcel.annotation.ExcelRequestBody;
import org.easyj.easyjexcel.converter.ExcelConverter;
import org.springframework.core.MethodParameter;
import org.springframework.web.bind.support.WebDataBinderFactory;
import org.springframework.web.context.request.NativeWebRequest;
import org.springframework.web.method.support.HandlerMethodArgumentResolver;
import org.springframework.web.method.support.ModelAndViewContainer;
import org.springframework.web.multipart.MultipartFile;
import org.springframework.web.multipart.MultipartHttpServletRequest;
import org.springframework.web.util.WebUtils;

public class ExcelRequestBodyHandler implements HandlerMethodArgumentResolver {

	private ExcelConverter converters;

	public void setConverters(ExcelConverter converters) {
		this.converters = converters;
	}

	@Override
	public boolean supportsParameter(MethodParameter parameter) {
		return parameter.hasParameterAnnotation(ExcelRequestBody.class);
	}

	@Override
	public Object resolveArgument(MethodParameter parameter, ModelAndViewContainer mavContainer,
			NativeWebRequest webRequest, WebDataBinderFactory binderFactory) throws Exception {

		HttpServletRequest servletRequest = webRequest.getNativeRequest(HttpServletRequest.class);
		MultipartHttpServletRequest multipartRequest = WebUtils.getNativeRequest(servletRequest,
				MultipartHttpServletRequest.class);

		ExcelRequestBody annotation = parameter.getParameterAnnotation(ExcelRequestBody.class);
		if (multipartRequest != null) {
			List<Object> result = new ArrayList<>();
			List<MultipartFile> files = multipartRequest.getFiles(annotation.name());
			for (MultipartFile file : files) {
				if (converters.supportsExcelType(annotation.type())) {
					List<?> part = converters.fromExcel(annotation, file.getInputStream());
					result.addAll(part);
				}
			}
			return result;
		}
		return null;

	}

}
