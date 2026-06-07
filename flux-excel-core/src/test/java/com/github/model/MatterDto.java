package com.github.model;

import com.github.excel.annotation.ExcelWrite;
import com.github.excel.annotation.ExcelWriteProperty;
import com.github.excel.enums.ExcelThemeEnum;
import com.github.excel.model.ExcelBaseModel;
import lombok.Builder;
import lombok.Data;

import java.util.Date;

/**
 * @Author: Vachel Wang
 * @Date: 2026/4/24
 * @Description: 测试demo
 */
@Data
@ExcelWrite(nameSpace = "matter",freezeTitle = true,filterTitle = true,incrementSequenceNo = false,incrementSequenceTitle = "编号",rowIndex = 0, colIndex = 0, themeName = ExcelThemeEnum.NONE,fillTitle = true)
@Builder
public class MatterDto extends ExcelBaseModel {
	@ExcelWriteProperty(titleName = "物料编码",colWidth = 200,rowHeight = 70)
	private String matterCode;
	@ExcelWriteProperty(titleName = "物料名称",rowHeight = 20)
	private String matterName;
	@ExcelWriteProperty(titleName = "规格/品牌",colWidth = 200)
	private String brand;
	@ExcelWriteProperty(titleName = "配置/材质/重量",colWidth = 200)
	private String weight;
	@ExcelWriteProperty(titleName = "单位",colWidth = 200)
	private String unit;
	@ExcelWriteProperty(titleName = "供应商",colWidth = 200)
	private String supplier;
	@ExcelWriteProperty(titleName = "上次采购价格",colWidth = 200)
	private Double prevPurchasePrice;
	@ExcelWriteProperty(titleName = "初始报价",colWidth = 200)
	private Double initialPrice;
	@ExcelWriteProperty(titleName = "最终成交价",colWidth = 200)
	private Double finalPrice;
	@ExcelWriteProperty(titleName = "税率",colWidth = 200)
	private Double rate;
	@ExcelWriteProperty(titleName = "交货期/采购提前期",colWidth = 200)
	private Integer deliveryDate;
	@ExcelWriteProperty(titleName = "运费承担",colWidth = 200)
	private String freight;
	@ExcelWriteProperty(titleName = "付款方式",colWidth = 200)
	private String payType;
	@ExcelWriteProperty(titleName = "联系方式",colWidth = 200)
	private String concatMobile;
	@ExcelWriteProperty(titleName = "联系人",colWidth = 200)
	private String concatName;
	@ExcelWriteProperty(titleName = "建议采购",colWidth = 200)
	private String suggest;
	@ExcelWriteProperty(titleName = "备注",colWidth = 200)
	private String remark;
	@ExcelWriteProperty(titleName = "询价时间",colWidth = 200,formatPattern = "yyyy-MM-dd")
	private Date queryTime;

}
