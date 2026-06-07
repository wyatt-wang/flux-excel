package com.github.excel.write.style;

import com.github.excel.enums.ExcelThemeEnum;
import com.github.excel.helper.ExcelHelper;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.DefaultIndexedColorMap;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.util.HashMap;
import java.util.Map;
import java.util.Objects;

/**
 * @Author: Vachel Wang
 * @Date: 2026/4/24
 * @Description: excel样式抽象类
 */
public abstract class AbstractExcelStyle {

	protected Workbook workbook;
	protected Map<String, CellStyle> styleMap;
	protected Map<String, Font> fontMap;
	protected Map<String, Color> colorMap ;

	public AbstractExcelStyle(Workbook workbook, Map<String, CellStyle> styleMap, Map<String, Font> fontMap, Map<String, Color> colorMap) {
		this.workbook = workbook;
		this.styleMap = styleMap;
		this.fontMap = fontMap;
		this.colorMap = colorMap;
	}

	/**
	 * 添加新样式
	 */
	public abstract void addNewStyle();

	/**
	 * 添加新字体
	 */
	public abstract void addNewFont();

	/**
	 * 添加新颜色
	 */
	public abstract void addNewColor();

	/**
	 * 创建样式
	 *
	 * @param name 名称
	 * @return CellStyle
	 */
	public CellStyle createStyle(String name) {
		CellStyle style = styleMap.get(name);
		if (Objects.nonNull(style)) {
			return style;
		}
		style = workbook.createCellStyle();
		styleMap.put(name, style);
		return style;
	}

	/**
	 * 创建字体
	 *
	 * @param name 名称
	 * @return Font
	 */
	public Font createFont(String name) {
		Font font = fontMap.get(name);
		if (Objects.nonNull(font)) {
			return font;
		}
		font = workbook.createFont();
		fontMap.put(name, font);
		return font;
	}

	/**
	 * 创建颜色
	 * @param name
	 * @param r
	 * @param g
	 * @param b
	 * @param a
	 * @param colorIndex
	 * @return
	 */
	public Color createColor(String name, int r, int g, int b, int a, short colorIndex) {
		Color color = colorMap.get(name);
		if (Objects.nonNull(color)) {
			return color;
		}
		if (workbook instanceof HSSFWorkbook) {
			color = ExcelHelper.setCustomColor((HSSFWorkbook) workbook, (byte) r, (byte)g, (byte)b,colorIndex);
		} else if (workbook instanceof XSSFWorkbook) {
			color = new XSSFColor(new java.awt.Color(r, g, b, a), new DefaultIndexedColorMap());
		}
		if (Objects.nonNull(color)) {
			colorMap.put(name, color);
		}
		return color;
	}

	public void createTheme(ExcelThemeEnum themeEnum) {
		if (workbook instanceof HSSFWorkbook) {
			// ================= theme =================
			short titleRowBackGroundColorIndex = ((HSSFColor)createColor(themeEnum.getTitleBackGroundColorName(),  53,92,183,255,HSSFColor.HSSFColorPredefined.GREEN.getIndex())).getIndex();
			short borderColorIndex = ((HSSFColor) createColor(themeEnum.getBorderColorName(), 125, 152, 210, 255, HSSFColor.HSSFColorPredefined.GREY_25_PERCENT.getIndex())).getIndex();
			short oDDColorIndex = ((HSSFColor) createColor(themeEnum.getOddRowBackGroundColorName(), 208, 219, 239, 255, HSSFColor.HSSFColorPredefined.BLUE_GREY.getIndex())).getIndex();
			short eventColorIndex = ((HSSFColor) createColor(themeEnum.getEventRowBackGroundColorName(), 255, 255, 255, 255, HSSFColor.HSSFColorPredefined.WHITE.getIndex())).getIndex();

			CellStyle style = this.createStyle(themeEnum.getTitleRowStyleName());
			style.setVerticalAlignment(VerticalAlignment.CENTER);
			style.setAlignment(HorizontalAlignment.CENTER);
			style.setFillForegroundColor(titleRowBackGroundColorIndex);
			style.setFillPattern(FillPatternType.SOLID_FOREGROUND);

			style.setBorderBottom(BorderStyle.THIN);
			style.setBorderLeft(BorderStyle.THIN);
			style.setBorderRight(BorderStyle.THIN);
			style.setBorderTop(BorderStyle.THIN);

			style.setBottomBorderColor(borderColorIndex);
			style.setTopBorderColor(borderColorIndex);
			style.setLeftBorderColor(borderColorIndex);
			style.setRightBorderColor(borderColorIndex);
			style.setFont(fontMap.get(themeEnum.getTitleFontName()));

			createHSSFContentTheme(borderColorIndex, oDDColorIndex,eventColorIndex, themeEnum);
			// ================= theme end=================
		} else if (workbook instanceof XSSFWorkbook || workbook instanceof SXSSFWorkbook) {
			XSSFColor titleRowBackGroundColor = (XSSFColor) createColor(themeEnum.getTitleBackGroundColorName(), 53, 92, 183, 255, (short) 0);
			XSSFColor borderColor =  (XSSFColor) createColor(themeEnum.getBorderColorName(),  125,152,210,255,(short)0);
			XSSFColor oDDColor = (XSSFColor) createColor(themeEnum.getOddRowBackGroundColorName(), 208, 219, 239,255, (short) 0);
			XSSFColor eventColor = (XSSFColor) createColor(themeEnum.getEventRowBackGroundColorName(), 255, 255, 255,255, (short) 0);

			XSSFCellStyle style = (XSSFCellStyle)this.createStyle(themeEnum.getTitleRowStyleName());
			style.setVerticalAlignment(VerticalAlignment.CENTER);
			style.setAlignment(HorizontalAlignment.CENTER);
			style.setFillForegroundColor(titleRowBackGroundColor);
			style.setFillPattern(FillPatternType.SOLID_FOREGROUND);

			style.setBorderBottom(BorderStyle.THIN);
			style.setBorderLeft(BorderStyle.THIN);
			style.setBorderRight(BorderStyle.THIN);
			style.setBorderTop(BorderStyle.THIN);

			style.setTopBorderColor(borderColor);
			style.setLeftBorderColor(borderColor);
			style.setRightBorderColor(borderColor);
			style.setBottomBorderColor(borderColor);

			style.setFont(fontMap.get(themeEnum.getTitleFontName()));

			createXSSFContentTheme(borderColor, oDDColor,eventColor, themeEnum);
		}
	}

	public void setDateFormat(CellStyle cellStyle, String formatStyleName) {
		short format = styleMap.get(formatStyleName).getDataFormat();
		cellStyle.setDataFormat(format);
	}

	public void createXSSFContentTheme(XSSFColor borderColor, XSSFColor oDDColor, XSSFColor eventColor, ExcelThemeEnum themeEnum) {
		XSSFCellStyle style;
		style = (XSSFCellStyle) this.createStyle(themeEnum.getEvenRowStyleName());
		style.setVerticalAlignment(VerticalAlignment.CENTER);
		style.setAlignment(HorizontalAlignment.LEFT);
		style.setFillForegroundColor(eventColor);
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		style.setWrapText(true);

		style.setBorderBottom(BorderStyle.THIN);
		style.setBorderLeft(BorderStyle.THIN);
		style.setBorderRight(BorderStyle.THIN);
		style.setBorderTop(BorderStyle.THIN);

		style.setBottomBorderColor(borderColor);
		style.setTopBorderColor(borderColor);
		style.setLeftBorderColor(borderColor);
		style.setRightBorderColor(borderColor);
		style.setFont(fontMap.get(themeEnum.getContentFontName()));

		style = (XSSFCellStyle) this.createStyle(themeEnum.getOddRowStyleName());
		style.setVerticalAlignment(VerticalAlignment.CENTER);
		style.setAlignment(HorizontalAlignment.LEFT);
		style.setFillForegroundColor(oDDColor);
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		style.setWrapText(true);

		style.setBorderBottom(BorderStyle.THIN);
		style.setBorderLeft(BorderStyle.THIN);
		style.setBorderRight(BorderStyle.THIN);
		style.setBorderTop(BorderStyle.THIN);

		style.setBottomBorderColor(borderColor);
		style.setTopBorderColor(borderColor);
		style.setLeftBorderColor(borderColor);
		style.setRightBorderColor(borderColor);
		style.setFont(fontMap.get(themeEnum.getContentFontName()));

		style = (XSSFCellStyle) this.createStyle(themeEnum.getEvenRowStyleDateName());
		style.setVerticalAlignment(VerticalAlignment.CENTER);
		style.setAlignment(HorizontalAlignment.LEFT);
		style.setFillForegroundColor(oDDColor);
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		style.setWrapText(true);

		style.setBorderBottom(BorderStyle.THIN);
		style.setBorderLeft(BorderStyle.THIN);
		style.setBorderRight(BorderStyle.THIN);
		style.setBorderTop(BorderStyle.THIN);

		style.setBottomBorderColor(borderColor);
		style.setTopBorderColor(borderColor);
		style.setLeftBorderColor(borderColor);
		style.setRightBorderColor(borderColor);
		style.setFont(fontMap.get(themeEnum.getContentFontName()));

	}

	public void createHSSFContentTheme(short borderColorIndex, short oDDColorIndex,short eventColorIndex, ExcelThemeEnum themeEnum) {
		CellStyle style;
		style = this.createStyle(themeEnum.getEvenRowStyleName());
		style.setVerticalAlignment(VerticalAlignment.CENTER);
		style.setAlignment(HorizontalAlignment.LEFT);
		style.setFillForegroundColor(eventColorIndex);
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		style.setWrapText(true);

		style.setBorderBottom(BorderStyle.THIN);
		style.setBorderLeft(BorderStyle.THIN);
		style.setBorderRight(BorderStyle.THIN);
		style.setBorderTop(BorderStyle.THIN);

		style.setBottomBorderColor(borderColorIndex);
		style.setTopBorderColor(borderColorIndex);
		style.setLeftBorderColor(borderColorIndex);
		style.setRightBorderColor(borderColorIndex);
		style.setFont(fontMap.get(themeEnum.getContentFontName()));

		style = this.createStyle(themeEnum.getOddRowStyleName());
		style.setVerticalAlignment(VerticalAlignment.CENTER);
		style.setAlignment(HorizontalAlignment.LEFT);
		style.setFillForegroundColor(oDDColorIndex);
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		style.setWrapText(true);

		style.setBorderBottom(BorderStyle.THIN);
		style.setBorderLeft(BorderStyle.THIN);
		style.setBorderRight(BorderStyle.THIN);
		style.setBorderTop(BorderStyle.THIN);

		style.setBottomBorderColor(borderColorIndex);
		style.setTopBorderColor(borderColorIndex);
		style.setLeftBorderColor(borderColorIndex);
		style.setRightBorderColor(borderColorIndex);
		style.setFont(fontMap.get(themeEnum.getContentFontName()));
	}

}

