package com.mycomponents.excelhandle.excel;

import java.beans.IntrospectionException;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;
import java.util.Set;

import org.apache.commons.beanutils.PropertyUtils;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * @author Administrator
 * @date 2017年7月28日
 */
public class ExcelUtils {

	public static HSSFWorkbook exportReflectInfo(Class<?> c, List<?> voList, String type)
			throws IntrospectionException, IllegalAccessException, InvocationTargetException, NoSuchMethodException {

		HSSFWorkbook wb = new HSSFWorkbook();
		HSSFSheet sheet = wb.createSheet(type);
		HSSFRow row = sheet.createRow((int) 0); // 表头

		HSSFCellStyle style = wb.createCellStyle();
		style.setAlignment(HSSFCellStyle.ALIGN_CENTER); // 创建一个居中格式

		HSSFCellStyle cellStyle = wb.createCellStyle(); // 创建一个样式
		cellStyle.setFillForegroundColor(HSSFColor.YELLOW.index); // 设置颜色为黄色
		cellStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
		cellStyle.setAlignment(HSSFCellStyle.ALIGN_CENTER); // 创建一个居中格式

		HSSFFont font = wb.createFont();
		font.setFontName("宋体");
		font.setFontHeightInPoints((short) 11);// 设置字体大小
		style.setFont(font);
		cellStyle.setFont(font);

		// 循环设置表头
		Map<Field, ExcelColumn> fieldAnnos = new HashMap<>();
		Field[] fields = c.getDeclaredFields();
		for (int i = 0; i < fields.length; i++) {
			Field field = fields[i];
			ExcelColumn excelColumns = field.getAnnotation(ExcelColumn.class);
			if (excelColumns != null) {
				HSSFCell cell = row.createCell(excelColumns.sort());
				sheet.setColumnWidth(i, excelColumns.columnWidth());
				cell.setCellValue(excelColumns.name());
				cell.setCellStyle(cellStyle);
				fieldAnnos.put(field, excelColumns);
			}
		}
		// 设置内容
		for (int i = 0; i < voList.size(); i++) {
			row = sheet.createRow(i + 1);
			row.setRowStyle(style);
			Object vo = voList.get(i);
			Set<Entry<Field, ExcelColumn>> fieldEntry = fieldAnnos.entrySet();
			for (Entry<Field, ExcelColumn> entry : fieldEntry) {
				Object result = PropertyUtils.getProperty(vo, entry.getKey().getName());
				ExcelColumn col = entry.getValue();
				if (result == null) {
					row.createCell(col.sort()).setCellValue("");
				} else if (result instanceof Date) {
					row.createCell(entry.getValue().sort())
							.setCellValue(new SimpleDateFormat("yyyy-MM-dd HH:mm:ss").format(result));
				} else {
					row.createCell(entry.getValue().sort()).setCellValue(String.valueOf(result));
				}
			}
		}

		return wb;
	}

	/**
	 * 导入Excel表
	 * 
	 * @param c
	 * @param fileName
	 * @return
	 * @throws IntrospectionException
	 * @throws IllegalAccessException
	 * @throws InvocationTargetException
	 * @throws NoSuchMethodException
	 * @throws FileNotFoundException
	 * @throws IOException
	 */
	public static List<Map<String, Object>> importReflectInfo(Class<?> c, InputStream input, String fileName)
			throws IntrospectionException, IllegalAccessException, InvocationTargetException, NoSuchMethodException,
			FileNotFoundException, IOException {
		// 建立输入流,根据文件格式(2003或者2007)来初始化
		Workbook wb = null;
		boolean isE2007 = false; // 判断是否是excel2007及以上格式
		if (fileName.endsWith("xlsx")) {
			isE2007 = true;
		}
		if (isE2007) {
			wb = new XSSFWorkbook(input);
		} else {
			wb = new HSSFWorkbook(input);
		}
		// 获得第一个表单
		Sheet sheet = wb.getSheetAt(0);
		wb.close();
		// 存储定义的表头字段
		Map<Field, ExcelColumn> fieldAnnos = new HashMap<>();
		Field[] fields = c.getDeclaredFields();
		for (int i = 0; i < fields.length; i++) {
			Field field = fields[i];
			ExcelColumn excelColumns = field.getAnnotation(ExcelColumn.class);
			if (excelColumns != null) {
				fieldAnnos.put(field, excelColumns);
			}
		}
		// 获取最大行数
		int totalRows = sheet.getLastRowNum();
		List<Map<String, Object>> voList = new ArrayList<>();
		for (int i = 1; i <= totalRows; i++) {
			// 从第二行开始取值
			Row row = sheet.getRow(i);
			Set<Entry<Field, ExcelColumn>> fieldEntry = fieldAnnos.entrySet();
			// 遍历定义的字段集合并逐行取值存voList
			Map<String, Object> vo = new HashMap<>();
			for (Entry<Field, ExcelColumn> entry : fieldEntry) {
				ExcelColumn col = entry.getValue();
				// 取出单元格内容
				Cell cell = row.getCell(col.sort());
				// 判断单元格格式
				if (cell != null) {
					switch (cell.getCellType()) {

					case HSSFCell.CELL_TYPE_NUMERIC: // 数字
						if (HSSFDateUtil.isCellDateFormatted(cell)) {
							vo.put(entry.getKey().getName(),
									HSSFDateUtil.getJavaDate(cell.getNumericCellValue()).toString());
							// 如果是date类型则 ，获取该cell的date值
						} else {
							vo.put(entry.getKey().getName(), cell.getNumericCellValue());
						}
						break;
					case HSSFCell.CELL_TYPE_STRING: // 字符串
						vo.put(entry.getKey().getName(), cell.getStringCellValue());
						break;
					case HSSFCell.CELL_TYPE_BOOLEAN: // Boolean
						vo.put(entry.getKey().getName(), cell.getBooleanCellValue());
						break;
					case HSSFCell.CELL_TYPE_FORMULA: // 公式
						vo.put(entry.getKey().getName(), cell.getCellFormula());
						break;
					case HSSFCell.CELL_TYPE_BLANK: // 空值
						vo.put(entry.getKey().getName(), null);
						break;
					case HSSFCell.CELL_TYPE_ERROR: // 故障
						vo.put(entry.getKey().getName(), null);
						break;
					default:
						vo.put(entry.getKey().getName(), null);
						break;
					}
				}
			}
			voList.add(vo);
		}
		return voList;
	}
}
