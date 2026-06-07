package com.github.excel.annotation;

import com.github.excel.constant.ExcelConstant;
import com.github.excel.enums.ExcelWriterCellTitleModelEnum;
import com.github.excel.enums.ExcelWriterFillStyleEnum;
import com.github.excel.write.ExcelDefaultWriterDataFormat;
import com.github.excel.write.ExcelWriterDataFormat;

import java.lang.annotation.*;


@Documented
@Retention(RetentionPolicy.RUNTIME)
@Target(ElementType.FIELD)
public @interface ExcelWriteProperty {
	/**
	 * 标题名称
	 * @return
	 */
	String titleName();

	/**
	 * 多级表头，兼容 fast-excel/EasyExcel 的 head/value 风格。
	 * 为空时使用 titleName。
	 */
	String[] head() default {};

	/**
	 * 多级表头别名，等价于 head。
	 */
	String[] value() default {};

	/**
	 * 标题填充样式
	 * @see ExcelWriterCellTitleModelEnum
	 */
	ExcelWriterCellTitleModelEnum titleModel() default ExcelWriterCellTitleModelEnum.DEFAULT;

	/**
	 * 标题样式名
	 * @return
	 */
	String titleStyleName() default ExcelConstant.NULL_STR;

	/**
	 * 内容样式名
	 * @return
	 */
	String contentStyleName() default ExcelConstant.NULL_STR;

	/**
	 * 填充样式
	 * @see ExcelWriterFillStyleEnum
	 */
	ExcelWriterFillStyleEnum fillStyle() default ExcelWriterFillStyleEnum.HORIZONTAL;

	/**
	 * 分隔符
	 * @return
	 */
	String separator() default ExcelConstant.NULL_STR;

	/**
	 * 行高 -2 默认为15px
	 * @return
	 */
	short rowHeight() default ExcelConstant.MINUS_TWO_SHORT;

	/**
	 * 行高（磅），优先级高于 rowHeight，-2 使用默认高度
	 * @return
	 */
	short rowHeightPoints() default ExcelConstant.MINUS_TWO_SHORT;

	/**
	 * 列宽 -2 默认 ， -1 自动列宽
	 * @return
	 */
	short colWidth() default ExcelConstant.MINUS_TWO_SHORT;

	/**
	 * 列宽（字符宽度），优先级高于 colWidth，-2 使用默认宽度，-1 自动列宽
	 * @return
	 */
	short colWidthChars() default ExcelConstant.MINUS_TWO_SHORT;

	/**
	 * 行号 -1 自动模式，跟随ExcelExport的走
	 * @return
	 */
	int rowIndex() default ExcelConstant.MINUS_ONE_SHORT;

	/**
	 * 列号 -1 自动模式，跟随ExcelExport的走
	 * @return
	 */
	int colIndex() default ExcelConstant.MINUS_ONE_SHORT;

	/**
	 * 链接名称
	 * @return
	 */
	String linkName() default ExcelConstant.NULL_STR;

	/**
	 * 动态链接展示名字段，字段值仍作为链接地址
	 * @return
	 */
	String linkNameField() default ExcelConstant.NULL_STR;

	/**
	 * 格式化字符串
	 * @return
	 */
	String formatPattern() default ExcelConstant.NULL_STR;

	/**
	 * 自定义格式化器
	 * @return
	 */
	Class<? extends ExcelWriterDataFormat> formatter() default ExcelDefaultWriterDataFormat.class;
	/**
	 * 合并行数量
	 * @return
	 */
	int mergeRowNum() default ExcelConstant.ZERO_SHORT;

	/**
	 * 标题列合并数量
	 * @return
	 */
	int mergeTitleColNum() default ExcelConstant.ZERO_SHORT;

	/**
	 * 内容列合并数量
	 * @return
	 */
	int mergeContentColNum() default ExcelConstant.ZERO_SHORT;

	/**
	 * 垂直模式下是否换行
	 * @return
	 */
	boolean verticalNewLine() default false ;

	/**
	 * 批注内容
	 * @return
	 */
	String commentText() default ExcelConstant.NULL_STR ;

	/**
	 * 批注字体
	 * @return
	 */
	String commentFontName() default ExcelConstant.NULL_STR ;

	/**
	 * 下拉选项
	 * @return
	 */
	String[] dropDownOptions() default {} ;

	/**
	 * 是否禁用
	 * @return
	 */
	boolean disable() default false;

	/**
	 * 排序号
	 * @return
	 */
	int index() default ExcelConstant.INT_10000;
}
