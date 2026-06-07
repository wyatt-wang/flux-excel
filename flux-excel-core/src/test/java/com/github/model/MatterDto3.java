package com.github.model;

import com.github.excel.annotation.ExcelWrite;
import com.github.excel.annotation.ExcelWriteProperty;
import com.github.excel.enums.ExcelWriterFillStyleEnum;
import com.github.excel.enums.ExcelWriterListFillTypeEnum;
import com.github.excel.model.ExcelBaseModel;
import com.github.excel.write.style.ExcelBasicStyle;
import lombok.Builder;
import lombok.Data;

/**
 * @Author: Vachel Wang
 * @Date: 2026/4/24
 * @Description: 测试demo
 */
@Data
@ExcelWrite(rowIndex = 2,fillTitle = true,contentStyleName = ExcelBasicStyle.STYLE_TITLE_RED_FONT,
		fillType = ExcelWriterListFillTypeEnum.SHIFT, colIndex = 0,fillStyle = ExcelWriterFillStyleEnum.VERTICAL,titleStyleName = ExcelBasicStyle.STYLE_LIST_TITLE)
@Builder
public class MatterDto3 extends ExcelBaseModel {

	@ExcelWriteProperty(titleName = "物料名称",colWidth = 200,rowHeight = -1)
	private String matterName;
	@ExcelWriteProperty(titleName = "物料编码",colWidth = 200,rowHeight = -1)
	private String matterCode;
	@ExcelWriteProperty(titleName = "品牌/材质/规格",colWidth = 200,rowHeight = -1)
	private String brand;
	@ExcelWriteProperty(titleName = "采购数量",colWidth = 200,rowHeight = -1)
	private String num;
	@ExcelWriteProperty(titleName = "单位",colWidth = 200,rowHeight = -1)
	private String unit;
}
