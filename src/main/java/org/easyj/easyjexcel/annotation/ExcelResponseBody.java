package org.easyj.easyjexcel.annotation;

import java.lang.annotation.Documented;
import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

import org.easyj.easyjexcel.constant.ExcelType;

@Target({ ElementType.METHOD, ElementType.TYPE })
@Retention(RetentionPolicy.RUNTIME)
@Documented
public @interface ExcelResponseBody {

	String name();

	String sheetName() default "";

	boolean hasSeq() default true;

	ExcelType type() default ExcelType.XLS;
}
