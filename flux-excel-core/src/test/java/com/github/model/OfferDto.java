package com.github.model;

import com.github.excel.annotation.ExcelWrite;
import com.github.excel.annotation.ExcelWriteProperty;
import com.github.excel.enums.ExcelWriterListFillTypeEnum;
import com.github.excel.enums.ExcelThemeEnum;
import com.github.excel.model.ExcelBaseModel;
import lombok.Builder;
import lombok.Data;


/**
 * @Author: Vachel Wang
 * @Date: 2026/4/24
 * @Description: 测试demo
 */
@Data
@ExcelWrite(fillType = ExcelWriterListFillTypeEnum.COVER,/*mergeContentRowNum = 1,mergeContentColNum = 1,mergeTitleColNum = 1,mergeTitleRowNum = 1, *//*fillStyle = ExcelExportFillStyleEnum.VERTICAL,titleStyleName = ExcelBasicStyle.STYLE_LIST_TITLE,contentStyleName = ExcelBasicStyle.STYLE_TITLE_RED_FONT,*/themeName = ExcelThemeEnum.ZEBRA)
@Builder
public class OfferDto extends ExcelBaseModel {
	@ExcelWriteProperty(titleName = "初始报价",colWidth = -1,rowHeight = -1)
	private Double initPrice;
	@ExcelWriteProperty(titleName = "最终报价",colWidth = -1,rowHeight = -1)
	private Double finalPrice;
	@ExcelWriteProperty(titleName = "净价",colWidth = -1,rowHeight = -1)
	private Double cleanPrice;
	@ExcelWriteProperty(titleName = "金额",colWidth = -1,rowHeight = -1)
	private Double price;
}
