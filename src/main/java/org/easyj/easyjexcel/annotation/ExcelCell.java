package org.easyj.easyjexcel.annotation;

import java.lang.annotation.Documented;
import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

@Target({ ElementType.FIELD })
@Retention(RetentionPolicy.RUNTIME)
@Documented
public @interface ExcelCell {

	/**
	 * @return Excel中的列（头）名
	 */
	String name();

	/**
	 * @return 列名对应的A,B,C,D...
	 */
	String columnLabel();

	/**
	 * @return 列名下标，优先级高于columnLabel
	 */
	int column() default -1;

	/**
	 * @return 是否导出数据
	 */
	boolean isExport() default true;

	/**
	 * @return 是否为重要字段（整列标红,着重显示）
	 */
	boolean isMark() default false;

	/**
	 * @return 枚举类，枚举类必须继承 ExcelEnum
	 */
	Class<?> enumClass() default void.class;

}