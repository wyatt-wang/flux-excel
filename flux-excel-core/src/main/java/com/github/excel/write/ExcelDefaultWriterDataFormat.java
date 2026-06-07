package com.github.excel.write;

import com.github.excel.constant.ExcelConstant;
import com.github.excel.util.StringUtil;
import lombok.extern.slf4j.Slf4j;

import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.LocalTime;
import java.time.Year;
import java.time.YearMonth;
import java.time.format.DateTimeFormatter;
import java.time.temporal.TemporalAccessor;
import java.util.Calendar;
import java.util.Date;

/**
 * @Author: Vachel Wang
 * @Date: 2026/4/24
 * @Description: 字段格式化
 */
@Slf4j
public class ExcelDefaultWriterDataFormat implements ExcelWriterDataFormat {

	@Override
	public Object format(Object data, String pattern) {
		if (data instanceof Date) {
			pattern = StringUtil.isEmpty(pattern) ? ExcelConstant.DEFAULT_DATE_FORMAT : pattern;
			SimpleDateFormat sdf = new SimpleDateFormat(pattern);
			data = sdf.format((Date) data);
		} else if (data instanceof Calendar) {
			pattern = StringUtil.isEmpty(pattern) ? ExcelConstant.DEFAULT_DATE_FORMAT : pattern;
			Calendar calendar = (Calendar) data;
			SimpleDateFormat sdf = new SimpleDateFormat(pattern);
			data = sdf.format(calendar.getTime());
		} else if (data instanceof TemporalAccessor) {
			pattern = StringUtil.isEmpty(pattern) ? getDefaultJavaTimePattern(data) : pattern;
			data = DateTimeFormatter.ofPattern(pattern).format((TemporalAccessor) data);
		} else if (data instanceof Number) {
			if (StringUtil.isEmpty(pattern)) {
				return data;
			}
			DecimalFormat df = new DecimalFormat(pattern);
			Number number = (Number)data;
			data = df.format(number.doubleValue());
		}

		return data;
	}

	private String getDefaultJavaTimePattern(Object data) {
		if (data instanceof LocalDateTime) {
			return ExcelConstant.DEFAULT_DATE_FORMAT;
		}
		if (data instanceof LocalDate) {
			return ExcelConstant.DEFAULT_DATE_DAY_FORMAT;
		}
		if (data instanceof LocalTime) {
			return ExcelConstant.DEFAULT_TIME_FORMAT;
		}
		if (data instanceof YearMonth) {
			return ExcelConstant.DEFAULT_YEAR_MONTH_FORMAT;
		}
		if (data instanceof Year) {
			return ExcelConstant.DEFAULT_YEAR_FORMAT;
		}
		return ExcelConstant.DEFAULT_DATE_FORMAT;
	}
}
