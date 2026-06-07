package com.github.excel.util;

import com.github.excel.constant.ExcelConstant;

/**
 * @description: 字符串工具类
 * @author: Vachel Wang
 * @create: 2019-09-07 16:28
 **/
public class StringUtil {
	public static boolean isEmpty(Object str) {
		return (str == null || ExcelConstant.NULL_STR.equals(str));
	}
	public static boolean notEmpty(Object str) {
		return !isEmpty(str);
	}

	public static String concatMethodName(String field , String prefix){
		char[] cs = field.toCharArray();
		if(cs[ExcelConstant.ZERO_SHORT]!='_' && cs[ExcelConstant.ZERO_SHORT]!='$') {
			cs[ExcelConstant.ZERO_SHORT] -= ExcelConstant.INT_32;
		}
		return prefix + String.valueOf(cs);
	}
}
