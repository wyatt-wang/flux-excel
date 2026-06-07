package com.github.model;

import com.github.excel.annotation.ExcelWrite;
import com.github.excel.annotation.ExcelWriteProperty;
import com.github.excel.enums.ExcelWriterFillStyleEnum;
import com.github.excel.enums.ExcelWriterListFillTypeEnum;
import com.github.excel.enums.ExcelThemeEnum;
import com.github.excel.model.ExcelBaseModel;
import com.github.excel.write.style.ExcelBasicStyle;
import lombok.Builder;
import lombok.Data;

import jakarta.validation.constraints.NotEmpty;

/**
 * @Author: Vachel Wang
 * @Date: 2026/4/24
 * @Description: 测试demo
 */
@Data
@ExcelWrite(rowIndex = 2,colIndex = 0,fillTitle = false,fillType = ExcelWriterListFillTypeEnum.SHIFT, fillStyle = ExcelWriterFillStyleEnum.VERTICAL,titleStyleName = ExcelBasicStyle.STYLE_ZEBRA_TITLE_ROW, themeName = ExcelThemeEnum.ZEBRA)
@Builder
public class MatterDto2 extends ExcelBaseModel {

	@ExcelWriteProperty(titleName = "物料名称")
	@NotEmpty
	private String matterName;
	@ExcelWriteProperty(titleName = "物料编码")
	private String matterCode;
	@ExcelWriteProperty(titleName = "品牌/材质/规格")
	private String brand;
	@ExcelWriteProperty(titleName = "采购数量")
	private Integer num;
	@ExcelWriteProperty(titleName = "单位")
	private String unit;
}
