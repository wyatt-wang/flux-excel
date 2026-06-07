package com.github.excel.model;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;
import lombok.experimental.Accessors;
import lombok.experimental.SuperBuilder;

/**
 * @Author: Vachel Wang
 * @Date: 2026/4/24
 * @Description: 读取图片
 */
@Data
@Accessors(chain = true)
@SuperBuilder
@AllArgsConstructor
@NoArgsConstructor
public class ExcelReaderPictureModel {
	/**
	 * sheet 名称
	 */
	private String sheetName;
	/**
	 * 行号
	 */
	private Integer rowIndex;
	/**
	 * 列号
	 */
	private Integer colIndex;
	/**
	 * 坐标
	 */
	private String point;
	/**
	 * 流
	 */
	private byte[] bytes;
	/**
	 * 后缀
	 */
	private String suffix;
}
