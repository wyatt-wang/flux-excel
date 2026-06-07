package com.github.excel.annotation;

import com.github.excel.constant.ExcelConstant;
import com.github.excel.enums.*;

import java.lang.annotation.*;


@Documented
@Inherited
@Retention(RetentionPolicy.RUNTIME)
@Target(ElementType.TYPE)
public @interface ExcelWrite {
	/**
	 * 命名空间
	 */
	String nameSpace() default ExcelConstant.NULL_STR;

	/**
	 * 作用域
	 * @see ExcelWriterScopeEnum
	 */
	ExcelWriterScopeEnum scope() default ExcelWriterScopeEnum.ALL_SHEET;

	/**
	 * 填充标题
	 */
	boolean fillTitle() default true;

	/**
	 * 冻结标题
	 * @return
	 */
	boolean freezeTitle() default false;

	/**
	 * 行号
	 * @return
	 */
	int rowIndex() default ExcelConstant.ZERO_SHORT;

	/**
	 * 列号
	 * @return
	 */
	int colIndex() default ExcelConstant.ZERO_SHORT;

	/**
	 * 填充样式
	 * @see ExcelWriterFillStyleEnum
	 */
	ExcelWriterFillStyleEnum fillStyle() default ExcelWriterFillStyleEnum.VERTICAL;

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
	 * 列表填充类型
	 * @see ExcelWriterListFillTypeEnum
	 */
	ExcelWriterListFillTypeEnum fillType() default ExcelWriterListFillTypeEnum.COVER;

	/**
	 * 标题行合并数量
	 * @return
	 */
	int mergeTitleRowNum() default ExcelConstant.ZERO_SHORT;

	/**
	 * 标题列合并数量
	 * @return
	 */
	int mergeTitleColNum() default ExcelConstant.ZERO_SHORT;

	/**
	 * 内容行合并数量
	 * @return
	 */
	int mergeContentRowNum() default ExcelConstant.ZERO_SHORT;

	/**
	 * 内容列合并数量
	 * @return
	 */
	int mergeContentColNum() default ExcelConstant.ZERO_SHORT;

	/**
	 * 标题填充样式
	 * @see ExcelWriterCellTitleModelEnum
	 */
	ExcelWriterCellTitleModelEnum titleModel() default ExcelWriterCellTitleModelEnum.DEFAULT;

	/**
	 * 兼容旧版枚举主题
	 * @return
	 */
	ExcelThemeEnum themeName() default ExcelThemeEnum.NONE;

	/**
	 * 自定义主题注册名
	 * @return
	 */
	String customThemeName() default "NONE";

	/**
	 * 是否填充自增序列号
	 * @return
	 */
	boolean incrementSequenceNo() default false;

	/**
	 * 是否填充自增序列号
	 * @return
	 */
	String incrementSequenceTitle() default ExcelConstant.INCREMENT_SEQUENCE_NO_TITLE_NAME;

	/**
	 * 是否填充内容
	 * @return
	 */
	boolean fillContent() default true;
	/**
	 * 筛选
	 * @return
	 */
	boolean filterTitle() default false;
}
