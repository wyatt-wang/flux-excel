package com.github.excel.write.style;

import com.github.excel.constant.ExcelConstant;
import com.github.excel.helper.ExcelHelper;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.DefaultIndexedColorMap;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.util.Map;
import java.util.Objects;

/**
 * @Author: Vachel Wang
 * @Date: 2026/4/24
 * @Description: Excel 内置样式
 */
public class ExcelBasicStyle extends AbstractExcelStyle {

	public static final String STYLE_FOREGROUND_COLOR_YELLOW = "BASIC_STYLE_FOREGROUND_COLOR_YELLOW";

	public static final String STYLE_AROUND_BORDER_READ = "BASIC_STYLE_AROUND_BORDER_READ";

	public static final String STYLE_TITLE = "BASIC_STYLE_TITLE";

	public static final String STYLE_LIST_TITLE = "BASIC_STYLE_LIST_TITLE";

	public static final String STYLE_TITLE_RED_FONT = "BASIC_STYLE_TITLE_RED_FONT";

	public static final String STYLE_CONTENT = "BASIC_STYLE_CONTENT";

	public static final String STYLE_HLINK = "BASIC_STYLE_HLINK";

	public static final String STYLE_ZEBRA_TITLE_ROW = "BASIC_STYLE_ZEBRA_TITLE";

	public static final String STYLE_ZEBRA_ODD_ROW = "BASIC_STYLE_ZEBRA_ODD_ROW";

	public static final String STYLE_ZEBRA_EVEN_ROW = "BASIC_STYLE_ZEBRA_EVEN_ROW";

	public static final String STYLE_ZEBRA_ODD_ROW_DATE = "BASIC_STYLE_ZEBRA_ODD_ROW_DATE";

	public static final String STYLE_ZEBRA_EVEN_ROW_DATE = "BASIC_STYLE_ZEBRA_EVEN_ROW_DATE";

	public static final String STYLE_DATE_YYYYMMDDHHMMSS = "BASIC_STYLE_DATE_YYYYMMDDHHMMSS";

	public static final String STYLE_DATE_YYYYMMDD = "BASIC_STYLE_DATE_YYYYMMDD";

	public static final String FONT_SIZE16_BLOLD = "BASIC_FONT_SIZE16_BLOLD";

	public static final String FONT_SIZE16_BLOLD_RED = "BASIC_FONT_SIZE16_BLOLD_RED";

	public static final String FONT_SIZE16_BLOLD_WHITE = "BASIC_FONT_SIZE16_BLOLD_WHITE";

	public static final String FONT_SIZE14 = "BASIC_FONT_SIZE14";

	public static final String FONT_SIZE14_BLOLD_UNDERLINE = "BASIC_FONT_SIZE14_BLOLD_UNDERLINE";

	public static final String FONT_HLINK = "BASIC_FONT_HLINK";

	public static final String COLOR_ZEBRA_TITLE_ROW_BACKGROUND = "COLOR_ZEBRA_TITLE_ROW_BACKGROUND";

	public static final String COLOR_ZEBRA_BORDER = "COLOR_ZEBRA_BORDER";

	public static final String COLOR_ZEBRA_ODD_ROW = "COLOR_ZEBRA_ODD_ROW";

	public static final String COLOR_ZEBRA_EVENT_ROW = "COLOR_ZEBRA_EVENT_ROW";

	private static final String DEFAULT_FONT_NAME = "宋体";


	public ExcelBasicStyle(Workbook workbook, Map<String, CellStyle> styleMap, Map<String, Font> fontMap, Map<String, Color> colorMap) {
		super(workbook, styleMap, fontMap, colorMap);
	}

	@Override
	public void addNewStyle() {
		CellStyle style = this.createStyle(STYLE_AROUND_BORDER_READ);
		style.setBorderBottom(BorderStyle.THIN);
		style.setBorderLeft(BorderStyle.THIN);
		style.setBorderRight(BorderStyle.THIN);
		style.setBorderTop(BorderStyle.THIN);

		style.setBottomBorderColor(IndexedColors.RED.index);
		style.setLeftBorderColor(IndexedColors.RED.index);
		style.setRightBorderColor(IndexedColors.RED.index);
		style.setTopBorderColor(IndexedColors.RED.index);

		style = this.createStyle(STYLE_FOREGROUND_COLOR_YELLOW);
		style.setFont(fontMap.get(FONT_SIZE14));
		style.setFillForegroundColor(IndexedColors.YELLOW.index);
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		style.setVerticalAlignment(VerticalAlignment.CENTER);
		style.setAlignment(HorizontalAlignment.LEFT);

		style = this.createStyle(STYLE_TITLE);
		style.setVerticalAlignment(VerticalAlignment.CENTER);
		style.setAlignment(HorizontalAlignment.RIGHT);
		style.setFont(fontMap.get(FONT_SIZE16_BLOLD));

		style = this.createStyle(STYLE_LIST_TITLE);
		style.setVerticalAlignment(VerticalAlignment.CENTER);
		style.setAlignment(HorizontalAlignment.CENTER);
		style.setFont(fontMap.get(FONT_SIZE16_BLOLD));

		style = this.createStyle(STYLE_TITLE_RED_FONT);
		style.setVerticalAlignment(VerticalAlignment.CENTER);
		style.setAlignment(HorizontalAlignment.CENTER);
		style.setFont(fontMap.get(FONT_SIZE16_BLOLD_RED));

		style = this.createStyle(STYLE_CONTENT);
		style.setVerticalAlignment(VerticalAlignment.CENTER);
		style.setAlignment(HorizontalAlignment.LEFT);
		style.setFont(fontMap.get(FONT_SIZE14));

		style = this.createStyle(STYLE_HLINK);
		style.setVerticalAlignment(VerticalAlignment.CENTER);
		style.setAlignment(HorizontalAlignment.LEFT);
		style.setFont(fontMap.get(FONT_HLINK));

		style = this.createStyle(STYLE_DATE_YYYYMMDDHHMMSS);
		CreationHelper creationHelper = workbook.getCreationHelper();
		style.setDataFormat(creationHelper.createDataFormat().getFormat(ExcelConstant.DEFAULT_DATE_FORMAT));


		style = this.createStyle(STYLE_DATE_YYYYMMDD);
		creationHelper = workbook.getCreationHelper();
		style.setDataFormat(creationHelper.createDataFormat().getFormat(ExcelConstant.DEFAULT_DATE_DAY_FORMAT));

	}

	@Override
	public void addNewFont() {
		Font font = this.createFont(FONT_SIZE16_BLOLD);
		font.setBold(true);
		font.setFontHeightInPoints((short) 16);
		font.setFontName(DEFAULT_FONT_NAME);

		font = this.createFont(FONT_SIZE16_BLOLD_RED);
		font.setBold(true);
		font.setFontHeightInPoints((short) 16);
		font.setFontName(DEFAULT_FONT_NAME);
		font.setColor(IndexedColors.RED.index);

		font = this.createFont(FONT_SIZE16_BLOLD_WHITE);
		font.setBold(true);
		font.setFontHeightInPoints((short) 16);
		font.setFontName(DEFAULT_FONT_NAME);
		font.setColor(IndexedColors.WHITE.index);

		font = this.createFont(FONT_SIZE14);
		font.setFontHeightInPoints((short) 14);
		font.setFontName(DEFAULT_FONT_NAME);

		font = this.createFont(FONT_SIZE14_BLOLD_UNDERLINE);
		font.setFontHeightInPoints((short) 14);
		font.setFontName(DEFAULT_FONT_NAME);
		font.setBold(true);
		font.setUnderline(FontUnderline.SINGLE.getByteValue());

		font = this.createFont(FONT_HLINK);
		font.setUnderline(Font.U_SINGLE);
		font.setColor(IndexedColors.BLUE.getIndex());


	}

	@Override
	public void addNewColor() {
		this.createColor(COLOR_ZEBRA_TITLE_ROW_BACKGROUND, 53, 92, 183, 255, HSSFColor.HSSFColorPredefined.GREEN.getIndex());
		this.createColor(COLOR_ZEBRA_BORDER, 125, 152, 210, 255, HSSFColor.HSSFColorPredefined.GREY_25_PERCENT.getIndex());
		this.createColor(COLOR_ZEBRA_ODD_ROW, 208, 219, 239, 255, HSSFColor.HSSFColorPredefined.BLUE_GREY.getIndex());
		this.createColor(COLOR_ZEBRA_EVENT_ROW, 255, 255, 255, 255, HSSFColor.HSSFColorPredefined.WHITE.getIndex());
	}
}
