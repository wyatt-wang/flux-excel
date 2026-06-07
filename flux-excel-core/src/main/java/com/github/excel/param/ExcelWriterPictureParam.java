package com.github.excel.param;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;
import lombok.experimental.Accessors;
import lombok.experimental.SuperBuilder;

import jakarta.validation.constraints.NotNull;

/**
 * @Author: Vachel Wang
 * @Date: 2026/4/24
 * @Description: 超链接model
 */
@Data
@Accessors(chain = true)
@SuperBuilder
@AllArgsConstructor
@NoArgsConstructor
public class ExcelWriterPictureParam extends ExcelWriterCellParam {
	/**
	 * @see org.apache.poi.ss.usermodel.Workbook
	 */
	@NotNull(message = "图片类型不能为空")
	private Integer pictureType;
}
