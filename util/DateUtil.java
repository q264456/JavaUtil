package com.tops.common.util;

import java.text.SimpleDateFormat;
import java.util.Date;

public class DateUtil {
	
	/**
	 * @param date 时间
	 * @param format 格式
	 * @return format格式的字符串
	 */
	public static String fmtDate(Date date, String format){
		SimpleDateFormat sdf = new SimpleDateFormat(format);
		return sdf.format(date);
	}

	/**
	 * @param date 时间
	 * @return yyyy-MM-dd HH:mm:ss
	 */
	public static String fmtDate(Date date){
		SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
		return sdf.format(date);
	}
	
	public static void main(String[] args) {
		Date date = new Date();
		System.out.println(date);
		System.out.println(DateUtil.fmtDate(date));
		System.out.println(DateUtil.fmtDate(date, "yyyy-MM-dd"));
	}
}
