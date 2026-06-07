package com.github.export;

import com.github.excel.write.style.AbstractExcelStyle;
import org.apache.poi.ss.usermodel.*;

import java.util.Map;

/**
 * @Author: Vachel Wang
 * @Date: 2026/4/24
 * @Description: Excel 自定义样式
 */
public class ExcelCustomStyle extends AbstractExcelStyle {
	public static final String STYLE_TITLE_TEST = "STYLE_TITLE_TEST";

	public static final String STYLE_CONTENT_TEST = "STYLE_CONTENT_TEST";

	public static final String FONT_SIZE16_BLOLD_RED_TEST = "FONT_SIZE16_BLOLD_RED_TEST";

	public static final String FONT_SIZE14_BLOLD_UNDERLINE_TEST = "FONT_SIZE14_BLOLD_UNDERLINE_TEST";

	private static final String DEFAULT_FONT_NAME1 = "黑体";


	public ExcelCustomStyle(Workbook workbook, Map<String, CellStyle> styleMap, Map<String, Font> fontMap, Map<String, Color> colorMap) {
		super(workbook, styleMap, fontMap, colorMap);
	}

	@Override
	public void addNewStyle() {
		CellStyle style = this.createStyle(STYLE_TITLE_TEST);
		style.setVerticalAlignment(VerticalAlignment.CENTER);
		style.setAlignment(HorizontalAlignment.RIGHT);
		style.setFont(fontMap.get(FONT_SIZE16_BLOLD_RED_TEST));
		style.setBorderBottom(BorderStyle.THIN);
		style.setBottomBorderColor(IndexedColors.BLACK.getIndex());
		style.setBorderLeft(BorderStyle.THIN);
		style.setLeftBorderColor(IndexedColors.GREEN.getIndex());
		style.setBorderRight(BorderStyle.THIN);
		style.setRightBorderColor(IndexedColors.BLUE.getIndex());
		style.setBorderTop(BorderStyle.MEDIUM_DASHED);
		style.setTopBorderColor(IndexedColors.BLACK.getIndex());

		style = this.createStyle(STYLE_CONTENT_TEST);
		style.setVerticalAlignment(VerticalAlignment.CENTER);
		style.setAlignment(HorizontalAlignment.CENTER);
		style.setFont(fontMap.get(FONT_SIZE14_BLOLD_UNDERLINE_TEST));

	}

	@Override
	public void addNewFont() {
		Font font = this.createFont(FONT_SIZE16_BLOLD_RED_TEST);
		font.setBold(true);
		font.setFontHeightInPoints((short) 16);
		font.setFontName(DEFAULT_FONT_NAME1);
		font.setColor(IndexedColors.RED.index);

		font = this.createFont(FONT_SIZE14_BLOLD_UNDERLINE_TEST);
		font.setFontHeightInPoints((short) 14);
		font.setFontName(DEFAULT_FONT_NAME1);
		font.setBold(true);
		font.setUnderline(FontUnderline.SINGLE.getByteValue());
	}

	@Override
	public void addNewColor() {
		// This test style does not define custom colors.
	}
}
