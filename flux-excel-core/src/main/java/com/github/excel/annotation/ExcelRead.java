package com.github.excel.annotation;

import java.lang.annotation.*;

/**
 * @Author: Vachel Wang
 * @Date: 2026/4/24
 * @Description: Excel 读取
 */
@Documented
@Inherited
@Retention(RetentionPolicy.RUNTIME)
@Target(ElementType.TYPE)
public @interface ExcelRead {
	/**
	 * 启用分隔符
	 */
	boolean enableSeparator() default false ;

	/**
	 * 校验表头
	 */
	boolean checkTitle() default false;
}
