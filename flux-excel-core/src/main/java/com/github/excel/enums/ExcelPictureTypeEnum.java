package com.github.excel.enums;

import com.google.common.collect.Maps;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.util.Map;
import java.util.Objects;

/**
 * @Author: Vachel Wang
 * @Date: 2026/4/24
 * @Description: Excel读取图片类型枚举
 */
public enum ExcelPictureTypeEnum {
	PICTURE_TYPE_EMF(XSSFWorkbook.PICTURE_TYPE_EMF,".EMF"),
	PICTURE_TYPE_WMF(XSSFWorkbook.PICTURE_TYPE_WMF,".WMF"),
	PICTURE_TYPE_PICT(XSSFWorkbook.PICTURE_TYPE_PICT,".PICT"),
	PICTURE_TYPE_JPEG(XSSFWorkbook.PICTURE_TYPE_JPEG,".JPEG"),
	PICTURE_TYPE_PNG(XSSFWorkbook.PICTURE_TYPE_PNG,".PNG"),
	PICTURE_TYPE_DIB(XSSFWorkbook.PICTURE_TYPE_DIB,".DIB"),
	PICTURE_TYPE_GIF(XSSFWorkbook.PICTURE_TYPE_GIF,".GIF"),
	PICTURE_TYPE_TIFF(XSSFWorkbook.PICTURE_TYPE_TIFF,".TIFF"),
	PICTURE_TYPE_EPS(XSSFWorkbook.PICTURE_TYPE_EPS,".EPS"),
	PICTURE_TYPE_BMP(XSSFWorkbook.PICTURE_TYPE_BMP,".BMP"),
	PICTURE_TYPE_WPG(XSSFWorkbook.PICTURE_TYPE_WPG,".WPG");

	private int type ;
	private String suffix ;

	ExcelPictureTypeEnum(int type, String suffix ){
		this.type = type;
		this.suffix = suffix;
	}

	private static final Map<Integer, ExcelPictureTypeEnum> typeEnumMap = Maps.newHashMap();
	static {
		for (ExcelPictureTypeEnum typeEnum : ExcelPictureTypeEnum.values()) {
			typeEnumMap.put(typeEnum.type, typeEnum);
		}
	}

	public static ExcelPictureTypeEnum getTypeEnum(int type) {
		return typeEnumMap.get(type);
	}

	public static String getTypeSuffix(int type) {
		ExcelPictureTypeEnum typeEnum = typeEnumMap.get(type);
		if (Objects.isNull(typeEnum)) {
			return null;
		}
		return typeEnum.suffix;
	}

}
