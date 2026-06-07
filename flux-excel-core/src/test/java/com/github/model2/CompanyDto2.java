package com.github.model2;

import com.github.excel.annotation.ExcelWrite;
import com.github.excel.annotation.ExcelWriteProperty;
import com.github.excel.enums.ExcelThemeEnum;
import com.github.excel.model.ExcelBaseModel;
import lombok.Data;

import java.util.Calendar;
import java.util.Date;

/**
 * @Author: Vachel Wang
 * @Date: 2026/4/24
 * @Description: 简单导出
 */
@Data
@ExcelWrite(themeName = ExcelThemeEnum.ZEBRA,incrementSequenceNo = true)
public class CompanyDto2 extends ExcelBaseModel {
	@ExcelWriteProperty(titleName = "公司名称")
	private String name;
	@ExcelWriteProperty(titleName = "公司地址")
	private String address;
	@ExcelWriteProperty(titleName = "公司创建时间")
	private Date createTime;
	@ExcelWriteProperty(titleName = "修改时间",index = 99)
	private Calendar updateTime;
	@ExcelWriteProperty(titleName = "公司人数")
	private Integer persons;
	@ExcelWriteProperty(titleName = "LOGO",index = 1)
	private byte[] logo;

}
