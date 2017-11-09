package org.easyj.easyjexcel.converter;

import java.io.InputStream;
import java.lang.reflect.Field;
import java.math.BigDecimal;
import java.text.SimpleDateFormat;
import java.util.AbstractList;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.stream.Collectors;
import java.util.stream.Stream;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.easyj.easyjexcel.annotation.ExcelCell;
import org.easyj.easyjexcel.annotation.ExcelRequestBody;
import org.easyj.easyjexcel.annotation.ExcelResponseBody;
import org.easyj.easyjexcel.constant.ExcelEnum;
import org.springframework.format.annotation.DateTimeFormat;
import org.springframework.format.annotation.DateTimeFormat.ISO;
import org.springframework.util.StringUtils;

public abstract class AbstractGenericPoiExcelConverter implements ExcelConverter {

	// excel中每个sheet中最多有65536行
	private static final int MAX_SHEET_SIZE = 65536;

	private static final ThreadLocal<Map<String, SimpleDateFormat>> simpleDateFormatCache = new ThreadLocal<>();

	@Override
	public List<?> fromExcel(ExcelRequestBody excelRequestBody, InputStream input) throws Exception {
		Workbook workBook = createWorkBook(input);
		return converterFrom(workBook, excelRequestBody.requireClass(), excelRequestBody.hasSeq());
	}

	@Override
	public <T> Workbook toExcel(ExcelResponseBody excelResponseBody, List<T> excelVoList) throws Exception {
		Workbook workBook = createWorkBook();
		return convertTo(excelVoList, workBook, excelResponseBody.hasSeq(), excelResponseBody.sheetName());
	}

	protected abstract Workbook createWorkBook(InputStream inputStream) throws Exception;

	protected abstract Workbook createWorkBook() throws Exception;

	private <T> List<T> converterFrom(Workbook book, Class<T> clazz, boolean hasSeq) throws Exception {
		List<T> result = new ArrayList<>();

		// 只处理第一个sheet的内容
		Sheet sheet = book.getSheetAt(0);

		// 得到数据的条目数，sheet.getLastRowNum()总是返回最后一条的索引
		int rows = sheet.getLastRowNum() + 1;

		// 只有表头则不处理
		if (rows <= 1) {
			return result;
		}

		// 得到类的所有field
		Field[] fieldArr = clazz.getDeclaredFields();
		// 定义一个map用于存放列的序号和field
		Map<Integer, Field> columnNo_field = new HashMap<>();
		for (Field field : fieldArr) {
			// 将有注解的field存放到map中
			if (field.isAnnotationPresent(ExcelCell.class)) {
				ExcelCell annotation = field.getAnnotation(ExcelCell.class);
				// 设置类的私有字段属性可访问
				field.setAccessible(true);

				columnNo_field.put(getColumnNo(annotation), field);
			}
		}
		// 从第2行开始取数据,默认第一行是表头
		for (int i = 1; i < rows; i++) {
			// 得到一行中的所有单元格对象.
			Row row = sheet.getRow(i);
			int cells = row.getLastCellNum() + 1;
			T entity = null;

			boolean havingSkip = false;
			for (int j = 0; j < cells; j++) {

				// 如果有序号列，则跳过
				if (hasSeq && !havingSkip) {
					havingSkip = true;
					continue;
				}

				// 从map中得到对应列的field
				Field field = columnNo_field.get(j);
				if (null == field) {
					continue;
				}
				// 获取单元格
				Cell cell = row.getCell(j);
				if (null == cell) {
					continue;
				}
				// 获取单元格中的内容.
				String value = cell.getStringCellValue();
				if (StringUtils.isEmpty(value)) {
					continue;
				}

				// 如果不存在实例则新建
				entity = (null == entity ? clazz.newInstance() : entity);

				resolveType(field, entity, value);

			}

			if (null != entity) {
				result.add(entity);
			}
		}
		return result;
	}

	private <T> void resolveType(Field field, T entity, String value) throws Exception {
		// 取得类型,并根据对象类型设置值.
		Class<?> fieldType = field.getType();

		ExcelCell annotation = field.getAnnotation(ExcelCell.class);
		if (annotation.enumClass() != null && annotation.enumClass() != void.class
				&& annotation.enumClass().isEnum() && !fieldType.isEnum()) {
			ExcelEnum<?>[] enumArr = (ExcelEnum<?>[]) annotation.enumClass().getEnumConstants();
			for (ExcelEnum<?> excelEnum : enumArr) {
				if (excelEnum.excelName().equalsIgnoreCase(value.trim())) {
					field.set(entity, excelEnum.excelValue());
					break;
				}
			}
		} else if (String.class == fieldType) {
			field.set(entity, value);
		} else if (BigDecimal.class == fieldType) {
			value = value.contains("%") ? value.replace("%", "") : value;
			field.set(entity, BigDecimal.valueOf(Double.valueOf(value)));
		} else if (Date.class == fieldType) {
			Date date = getSimpleDateFormat(field).parse(value);
			field.set(entity, date);
		} else if ((Integer.TYPE == fieldType) || (Integer.class == fieldType)) {
			field.set(entity, Integer.parseInt(value));
		} else if ((Long.TYPE == fieldType) || (Long.class == fieldType)) {
			field.set(entity, Long.valueOf(value));
		} else if ((Float.TYPE == fieldType) || (Float.class == fieldType)) {
			field.set(entity, Float.valueOf(value));
		} else if ((Short.TYPE == fieldType) || (Short.class == fieldType)) {
			field.set(entity, Short.valueOf(value));
		} else if ((Double.TYPE == fieldType) || (Double.class == fieldType)) {
			field.set(entity, Double.valueOf(value));
		} else if (Character.TYPE == fieldType) {
			field.set(entity, value.charAt(0));
		} else if (fieldType.isEnum()) {
			if (!ExcelEnum.class.isAssignableFrom(fieldType)) {
				Enum<?>[] enumArr = (Enum<?>[]) fieldType.getEnumConstants();
				for (Enum<?> enumObj : enumArr) {
					if (enumObj.name().equalsIgnoreCase(value.trim())) {
						field.set(entity, enumObj);
						break;
					}
				}
			} else {
				ExcelEnum<?>[] enumArr = (ExcelEnum<?>[]) fieldType.getEnumConstants();
				for (ExcelEnum<?> excelEnum : enumArr) {
					if (excelEnum.excelName().equalsIgnoreCase(value.trim())) {
						field.set(entity, excelEnum);
						break;
					}
				}
			}
		}
	}

	private <T> Workbook convertTo(List<T> excelVoList, Workbook workbook, boolean hasSeq, String sheetName)
			throws Exception {

		// 得到所有定义字段
		List<Field> fieldList = Stream.of(excelVoList.get(0).getClass().getDeclaredFields())
				.filter(field -> field.isAnnotationPresent(ExcelCell.class)).collect(Collectors.toList());

		// 普通表头样式
		Font normalHeadFont = workbook.createFont();
		normalHeadFont.setFontName("Arail narrow"); // 字体
		normalHeadFont.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD); // 字体宽度
		normalHeadFont.setColor(HSSFFont.COLOR_NORMAL); // 字体颜色
		CellStyle normalHeadCellStyle = workbook.createCellStyle();
		normalHeadCellStyle.setFillForegroundColor(IndexedColors.LIGHT_GREEN.index);
		normalHeadCellStyle.setFillPattern(CellStyle.SOLID_FOREGROUND);
		normalHeadCellStyle.setFont(normalHeadFont);

		// 标红表头样式
		Font markedHeadFont = workbook.createFont();
		markedHeadFont.setFontName("Arail narrow"); // 字体
		markedHeadFont.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD); // 字体宽度
		markedHeadFont.setColor(HSSFFont.COLOR_RED); // 字体颜色
		CellStyle markedHeadCellStyle = workbook.createCellStyle();
		markedHeadCellStyle.setFillForegroundColor(IndexedColors.LIGHT_GREEN.index);
		markedHeadCellStyle.setFillPattern(CellStyle.SOLID_FOREGROUND);
		markedHeadCellStyle.setFont(markedHeadFont);

		// 普通内容样式
		Font normalContextFont = workbook.createFont();
		normalContextFont.setColor(HSSFFont.COLOR_NORMAL); // 字体颜色
		CellStyle normalContextCellStyle = workbook.createCellStyle();
		normalContextCellStyle.setFont(normalContextFont);

		// 标红内容样式
		Font markedContextFont = workbook.createFont();
		CellStyle markedContextCellStyle = workbook.createCellStyle();
		markedContextFont.setColor(HSSFFont.COLOR_RED); // 字体颜色
		markedContextCellStyle.setFont(markedContextFont);

		// 计算一共有多少个sheet
		List<List<T>> partitionList = new Partition<>(excelVoList, MAX_SHEET_SIZE);

		for (int index = 0; index < partitionList.size(); index++) {
			// 产生工作表对象
			Sheet sheet = workbook.createSheet();
			// 设置工作表的名称
			workbook.setSheetName(index, sheetName + (index + 1));
			// 生成表头行
			Row headRow = sheet.createRow(0);

			// 生成表头的标志
			boolean havingHead = false;

			// 生成序号列
			if (hasSeq) {
				Cell cell = headRow.createCell(0);
				cell.setCellStyle(normalHeadCellStyle);
				cell.setCellValue("序号");
			}

			List<T> onePieceOfList = partitionList.get(index);
			// 生成内容列
			// 写入各条记录,每条记录对应excel表中的一行
			for (int i = 0; i < onePieceOfList.size(); i++) {

				Row contextRow = sheet.createRow(i + 1);

				if (hasSeq) {
					Cell cell = contextRow.createCell(0);
					cell.setCellStyle(normalContextCellStyle);
					cell.setCellValue(i + 1);
				}

				T vo = onePieceOfList.get(i); // 得到导出对象.

				for (Field field : fieldList) {
					// 获得field
					// 设置实体类私有属性可访问
					field.setAccessible(true);
					ExcelCell annotation = field.getAnnotation(ExcelCell.class);

					int columnNo = getColumnNo(annotation);

					if (!havingHead) {
						// 创建列
						Cell cell = headRow.createCell(columnNo);
						if (annotation.isMark()) {
							cell.setCellStyle(markedHeadCellStyle);
						} else {
							cell.setCellStyle(normalHeadCellStyle);
						}
						sheet.setColumnWidth(i, computeColumnWidth(annotation.name()));
						// 设置列中写入内容为String类型
						cell.setCellType(HSSFCell.CELL_TYPE_STRING);
						// 写入列名
						cell.setCellValue(annotation.name());
					}

					// 根据ExcelAttribute中设置情况决定是否导出,有些情况需要保持为空,希望用户填写这一列.
					if (annotation.isExport()) {
						// 创建cell
						Cell cell = contextRow.createCell(columnNo);
						if (annotation.isMark()) {
							cell.setCellStyle(markedContextCellStyle);
						} else {
							cell.setCellStyle(normalContextCellStyle);
						}
						// 如果数据存在就填入,不存在填入空格
						Object o = field.get(vo);
						String value = toObjectString(o, field, annotation);
						cell.setCellValue(value);
					}
				}
				// 仅第一遍生成表头
				havingHead = true;
			}
		}
		return workbook;
	}

	private int computeColumnWidth(String value) {
		return (int) ((value.getBytes().length <= 4 ? 6 : value.getBytes().length) * 1.5 * 256);
	}

	private String toObjectString(Object o, Field field, ExcelCell annotation) {
		if (null == o) {
			return "";
		}

		if (field.getType().isEnum()) {
			if (o instanceof ExcelEnum) {
				return ((ExcelEnum<?>) o).excelName();
			} else {
				return ((Enum<?>) o).name();
			}
		} else if (annotation.enumClass() != null && annotation.enumClass() != void.class
				&& annotation.enumClass().isEnum()) {
			ExcelEnum<?>[] enumArr = (ExcelEnum<?>[]) annotation.enumClass().getEnumConstants();
			for (ExcelEnum<?> excelEnum : enumArr) {
				if (excelEnum.excelValue().equals(o)) {
					return excelEnum.excelName();
				}
			}
		}

		if (o instanceof Date) {
			return getSimpleDateFormat(field).format((Date) o);
		}
		return String.valueOf(o);
	}

	private SimpleDateFormat getSimpleDateFormat(Field field) {
		DateTimeFormat annotation = field.getAnnotation(DateTimeFormat.class);
		String pattern = "yyyy-MM-dd HH:mm:ss";
		if (annotation != null) {
			if (!StringUtils.isEmpty(annotation.pattern())) {
				pattern = annotation.pattern();
			} else if (annotation.iso() != null && annotation.iso() != ISO.NONE) {
				switch (annotation.iso()) {
				case DATE:
					pattern = "yyyy-MM-dd";
					break;
				case TIME:
					pattern = "HH:mm:ss";
					break;
				default:
					break;
				}
			}
		}

		Map<String, SimpleDateFormat> simpleDateFormatMap = simpleDateFormatCache.get();
		if (simpleDateFormatMap == null) {
			simpleDateFormatMap = new HashMap<>();
			simpleDateFormatCache.set(simpleDateFormatMap);
		}

		return simpleDateFormatMap.computeIfAbsent(pattern, SimpleDateFormat::new);
	}

	/**
	 * 获取Excel列下标
	 */
	private int getColumnNo(ExcelCell annotation) {
		int columnNo = annotation.column();
		if (columnNo == -1) {
			String column = annotation.columnLabel();
			column = column.toUpperCase();

			// 从-1开始计算,字母重1开始运算。这种总数下来算数正好相同。
			int count = -1;
			char[] cs = column.toCharArray();
			for (int i = 0; i < cs.length; i++) {
				count += (cs[i] - 64) * Math.pow(26, cs.length - 1 - i);
			}

			columnNo = count;
		}

		return columnNo;
	}

	private static class Partition<T> extends AbstractList<List<T>> {
		private final List<T> list;
		private final int size;

		private Partition(final List<T> list, final int size) {
			this.list = list;
			this.size = size;
		}

		@Override
		public List<T> get(final int index) {
			final int listSize = size();
			if (listSize < 0) {
				throw new IllegalArgumentException("negative size: " + listSize);
			}
			if (index < 0) {
				throw new IndexOutOfBoundsException("Index " + index + " must not be negative");
			}
			if (index >= listSize) {
				throw new IndexOutOfBoundsException("Index " + index + " must be less than size " + listSize);
			}
			final int start = index * size;
			final int end = Math.min(start + size, list.size());
			return list.subList(start, end);
		}

		@Override
		public int size() {
			return (list.size() + size - 1) / size;
		}

		@Override
		public boolean isEmpty() {
			return list.isEmpty();
		}
	}

}