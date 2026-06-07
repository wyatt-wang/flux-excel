package com.github.model;

import com.github.excel.annotation.ExcelRead;
import com.github.excel.annotation.ExcelReadProperty;
import com.github.excel.model.ExcelBaseModel;
import lombok.Data;
import lombok.ToString;

/**
 * @Author: Vachel Wang
 * @Date: 2026/4/24
 * @Description: 简单导出
 */
@Data
@ExcelRead(enableSeparator = false)
@ToString
public class ProjectBidsBean extends ExcelBaseModel {
	@ExcelReadProperty(titleName = "规划标题")
	private String title;
	@ExcelReadProperty(titleName = "招标类别")
	private String type;
	@ExcelReadProperty(titleName = "招标模式")
	private String model;
	@ExcelReadProperty(titleName = "招标方式")
	private String bidsMethod;
	@ExcelReadProperty(titleName = "包干方式")
	private String doneMethod;
	@ExcelReadProperty(titleName = "标段数量")
	private Integer num;
	@ExcelReadProperty(titleName = "工期(天)")
	private Integer workTotalDays;
}
