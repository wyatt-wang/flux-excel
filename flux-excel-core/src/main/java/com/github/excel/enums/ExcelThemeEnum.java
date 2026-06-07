package com.github.excel.enums;

import com.github.excel.write.ExcelTheme;
import com.github.excel.write.style.ExcelBasicStyle;
import lombok.AllArgsConstructor;
import lombok.Getter;

/**
 * @Author: Vachel Wang
 * @Date: 2026/4/24
 * @Description: excel 主题样式
 */
@Getter
@AllArgsConstructor
public enum ExcelThemeEnum implements ExcelTheme {
	NONE(null, null, null, null, null, null, null, null, null, null, null, null, null, null),
	ZEBRA(ExcelBasicStyle.STYLE_ZEBRA_TITLE_ROW, ExcelBasicStyle.STYLE_ZEBRA_ODD_ROW, ExcelBasicStyle.STYLE_ZEBRA_EVEN_ROW,
			ExcelBasicStyle.STYLE_ZEBRA_ODD_ROW_DATE, ExcelBasicStyle.STYLE_ZEBRA_ODD_ROW_DATE, ExcelBasicStyle.COLOR_ZEBRA_TITLE_ROW_BACKGROUND,
			ExcelBasicStyle.COLOR_ZEBRA_BORDER, ExcelBasicStyle.COLOR_ZEBRA_ODD_ROW, ExcelBasicStyle.COLOR_ZEBRA_EVENT_ROW, ExcelBasicStyle.FONT_SIZE16_BLOLD_WHITE,
			ExcelBasicStyle.FONT_SIZE14, (short) 40, (short) 32, (short) 165),
	;
	/**
	 * 标题行样式名称
	 */
	private String titleRowStyleName;
	/**
	 * 奇数行样式名称
	 */
	private String oddRowStyleName;
	/**
	 * 偶数行样式名称
	 */
	private String evenRowStyleName;
	/**
	 * 奇数行时间类型样式名称
	 */
	private String oddRowStyleDateName;
	/**
	 * 偶数行时间类型样式名称
	 */
	private String evenRowStyleDateName;
	/**
	 * 标题背景颜色名称
	 */
	private String titleBackGroundColorName;
	/**
	 * 边框颜色名称
	 */
	private String borderColorName;
	/**
	 * 基数行颜色名称
	 */
	private String oddRowBackGroundColorName;
	/**
	 * 偶数行颜色名称
	 */
	private String eventRowBackGroundColorName;
	/**
	 * 标题字体名称
	 */
	private String titleFontName;
	/**
	 * 内容字体名称
	 */
	private String contentFontName;
	/**
	 * 标题默认高度
	 */
	private Short titleRowHeight;
	/**
	 * 内容默认高度
	 */
	private Short contentRowHeight;
	/**
	 * 单元格默认宽度
	 */
	private Short colWidth;


}
