package com.github.excel.read.format;

import com.github.excel.constant.ExcelConstant;
import com.github.excel.model.ExcelReaderPictureModel;
import com.github.excel.util.StringUtil;
import lombok.extern.slf4j.Slf4j;

import java.math.BigDecimal;
import java.math.BigInteger;
import java.text.DecimalFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.LocalTime;
import java.time.Year;
import java.time.YearMonth;
import java.time.format.DateTimeFormatter;
import java.util.Calendar;
import java.util.Date;
import java.util.Objects;

/**
 * @Author: Vachel Wang
 * @Date: 2026/4/24
 * @Description: 读取字段格式化
 */
@Slf4j
public class ExcelDefaultReaderDataFormat implements ExcelReaderDataFormat {

	@Override
	public Object format(Object data, String pattern, Class<?> targetCla) throws ParseException {
		Object result = data;
		if (Objects.isNull(result) || result instanceof ExcelReaderPictureModel) {
			return result;
		}
		if (targetCla == Date.class) {
			if (data instanceof String) {
				if (StringUtil.isEmpty(pattern)) {
					pattern = ExcelConstant.DEFAULT_DATE_FORMAT;
				}
				SimpleDateFormat sdf = new SimpleDateFormat(pattern);
				result = sdf.parse((String) data);
			}
		}else if(targetCla == LocalDateTime.class){
			if (data instanceof String) {
				if (StringUtil.isEmpty(pattern)) {
					pattern = ExcelConstant.DEFAULT_DATE_FORMAT;
				}
				result = LocalDateTime.parse(data.toString(), DateTimeFormatter.ofPattern(pattern));
			}
		}else if(targetCla == LocalDate.class){
			if (data instanceof String) {
				if (StringUtil.isEmpty(pattern)) {
					pattern = ExcelConstant.DEFAULT_DATE_DAY_FORMAT;
				}
				result = LocalDate.parse(data.toString(), DateTimeFormatter.ofPattern(pattern));
			}
		}else if(targetCla == YearMonth.class){
			if (data instanceof String) {
				if (StringUtil.isEmpty(pattern)) {
					pattern = ExcelConstant.DEFAULT_YEAR_MONTH_FORMAT;
				}
				result = YearMonth.parse(data.toString(), DateTimeFormatter.ofPattern(pattern));
			}
		}else if(targetCla == LocalTime.class){
			if (data instanceof String) {
				if (StringUtil.isEmpty(pattern)) {
					pattern = ExcelConstant.DEFAULT_TIME_FORMAT;
				}
				result = LocalTime.parse(data.toString(), DateTimeFormatter.ofPattern(pattern));
			}
		}else if(targetCla == Year.class){
			if (data instanceof String) {
				if (StringUtil.isEmpty(pattern)) {
					pattern = ExcelConstant.DEFAULT_YEAR_FORMAT;
				}
				result = Year.parse(data.toString(), DateTimeFormatter.ofPattern(pattern));
			}
		} else if (targetCla == Calendar.class) {
			if (data instanceof String) {
				if (StringUtil.isEmpty(pattern)) {
					pattern = ExcelConstant.DEFAULT_DATE_FORMAT;
				}
				SimpleDateFormat sdf = new SimpleDateFormat(pattern);
				Date date = sdf.parse((String) data);
				Calendar calendar = Calendar.getInstance();
				calendar.setTime(date);
				result = calendar;
			}
		} else if (Number.class.isAssignableFrom(targetCla)) {
			DecimalFormat df = StringUtil.isEmpty(pattern) ? new DecimalFormat() : new DecimalFormat(pattern);
			Number number = df.parse(data.toString());
			if (targetCla == Integer.class) {
				result = number.intValue();
			} else if (targetCla == Long.class) {
				result = number.longValue();
			} else if (targetCla == BigInteger.class) {
				result = new BigInteger(String.valueOf(number.longValue()));
			} else if (targetCla == Short.class) {
				result = number.shortValue();
			} else if (targetCla == Byte.class) {
				result = number.byteValue();
			} else if (targetCla == Float.class) {
				result = number.floatValue();
			} else if (targetCla == Double.class) {
				result = number.doubleValue();
			} else if (targetCla == BigDecimal.class) {
				result = new BigDecimal(String.valueOf(number.doubleValue()));
			}
		} else if (targetCla == Boolean.class) {
			if (data instanceof String) {
				result = Boolean.valueOf((String) data);
			}
		}
		return result;
	}
}
