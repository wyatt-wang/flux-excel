package com.github.excel.annotation;

import com.github.excel.constant.ExcelConstant;
import com.github.excel.read.format.ExcelDefaultReaderDataFormat;
import com.github.excel.read.format.ExcelReaderDataFormat;

import java.lang.annotation.*;

/**
 * @Author: Vachel Wang
 * @Date: 2026/4/24
 * @Description: Excel读取字段
 */
@Documented
@Inherited
@Retention(RetentionPolicy.RUNTIME)
@Target(ElementType.FIELD)
public @interface ExcelReadProperty {
	/**
	 * 标题名称
	 */
	String titleName() default ExcelConstant.NULL_STR;
	/**
	 * 分隔符
	 * @return
	 */
	String separator() default ExcelConstant.COLON;
	/**
	 * 格式化字符串
	 * @return
	 */
	String formatPattern() default ExcelConstant.NULL_STR;
	/**
	 * 自定义格式化器
	 * @return
	 */
	Class<? extends ExcelReaderDataFormat> formatter() default ExcelDefaultReaderDataFormat.class;
	/**
	 * 是否校验空数据
	 */
	boolean checkNull() default false;
	/**
	 * 是否禁用
	 * @return
	 */
	boolean disable() default false;
}
