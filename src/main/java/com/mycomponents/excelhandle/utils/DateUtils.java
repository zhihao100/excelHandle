package com.mycomponents.excelhandle.utils;

import java.text.DateFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Date;

public class DateUtils {
	/**
	 * 格式化(yyyy-MM-dd HH:mm:ss)过的时间类型转Date
	 *
	 * @param time
	 * @return: Date
	 * @throws Exception
	 */
	public static Date stringParseToDate(String time) throws Exception {
		DateFormat format = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
		Date date = null;
		try {
			date = format.parse(time);
		} catch (ParseException e) {
			throw new Exception("转换日期失败", new Throwable(e));
		}
		return date;
	}

	/**
	 * 未格式化过的时间类型Date转yyyy-MM-dd HH:mm:ss样式
	 *
	 * @param time
	 * @return: String
	 * @throws Exception
	 */
	public static String dateParseToString(Date time) throws Exception {
		SimpleDateFormat dateFormater = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
		return dateFormater.format(time);
	}
}
