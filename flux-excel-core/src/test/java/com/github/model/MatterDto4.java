package com.github.model;

import com.github.excel.annotation.ExcelWrite;
import com.github.excel.annotation.ExcelWriteProperty;
import com.github.excel.enums.ExcelWriterCellTitleModelEnum;
import com.github.excel.enums.ExcelWriterFillStyleEnum;
import com.github.excel.enums.ExcelThemeEnum;
import com.github.excel.model.ExcelBaseModel;
import com.github.excel.write.style.ExcelBasicStyle;
import com.github.export.ExcelCustomStyle;
import lombok.Data;

import java.util.Date;

/**
 * @Author: Vachel Wang
 * @Date: 2026/4/24
 * @Description: 测试demo
 */
@Data
@ExcelWrite(nameSpace = "matter",rowIndex = 2, colIndex = 3,mergeTitleRowNum = 2,mergeContentRowNum = 2/*,titleStyleName = ExcelBasicStyle.STYLE_TITLE_RED_FONT,contentStyleName = ExcelBasicStyle.STYLE_CONTENT*/,fillStyle = ExcelWriterFillStyleEnum.HORIZONTAL,titleModel = ExcelWriterCellTitleModelEnum.STAND_ALONE, themeName = ExcelThemeEnum.ZEBRA)
public class MatterDto4 extends ExcelBaseModel {
	@ExcelWriteProperty(titleName = "物料编码",colWidth = 200,commentText = "物料编码要正确填写哦！",commentFontName = ExcelBasicStyle.FONT_HLINK)
	private String matterCode;
	@ExcelWriteProperty(titleName = "物料名称", /*contentStyleName = ExcelBasicStyle.STYLE_TITLE_RED_FONT,*/colWidth = 200)
	private String matterName;
	@ExcelWriteProperty(titleName = "规格/品牌",colWidth = 200,fillStyle = ExcelWriterFillStyleEnum.VERTICAL,titleStyleName = ExcelCustomStyle.STYLE_TITLE_TEST)
	private String brand;
	@ExcelWriteProperty(titleName = "配置/材质/重量",colWidth = 300,rowHeight = 80)
	private String weight;
	@ExcelWriteProperty(titleName = "单位",colWidth = 200,fillStyle = ExcelWriterFillStyleEnum.VERTICAL,rowHeight = 80)
	private String unit;
	@ExcelWriteProperty(titleName = "供应商",colWidth = 200)
	private String supplier;
	@ExcelWriteProperty(titleName = "上次采购价格",colWidth = 200,fillStyle = ExcelWriterFillStyleEnum.VERTICAL)
	private Double prevPurchasePrice;
	@ExcelWriteProperty(titleName = "初始报价",colWidth = 200)
	private Double initialPrice;
	@ExcelWriteProperty(titleName = "最终成交价",colWidth = 200,fillStyle = ExcelWriterFillStyleEnum.VERTICAL)
	private Double finalPrice;
	@ExcelWriteProperty(titleName = "税率",colWidth = 200)
	private Double rate;
	@ExcelWriteProperty(titleName = "交货期/采购提前期",colWidth = 200,fillStyle = ExcelWriterFillStyleEnum.VERTICAL)
	private Integer deliveryDate;
	@ExcelWriteProperty(titleName = "运费承担",colWidth = 200)
	private String freight;
	@ExcelWriteProperty(titleName = "付款方式",colWidth = 200,fillStyle = ExcelWriterFillStyleEnum.VERTICAL)
	private String payType;
	@ExcelWriteProperty(titleName = "联系方式",colWidth = 200)
	private String concatMobile;
	@ExcelWriteProperty(titleName = "联系人",colWidth = 200,fillStyle = ExcelWriterFillStyleEnum.VERTICAL)
	private String concatName;
	@ExcelWriteProperty(titleName = "建议采购",colWidth = 200)
	private String suggest;
	@ExcelWriteProperty(titleName = "备注",colWidth = 300,fillStyle = ExcelWriterFillStyleEnum.VERTICAL)
	private String remark;
	@ExcelWriteProperty(titleName = "询价时间",colWidth = 200,formatPattern = "yyyy-MM-dd",commentText = "询价时间要正确填写哦！",commentFontName = ExcelBasicStyle.FONT_HLINK)
	private Date queryTime;

}
