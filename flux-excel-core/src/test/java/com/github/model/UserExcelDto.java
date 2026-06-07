package com.github.model;

import com.github.excel.annotation.ExcelWrite;
import com.github.excel.annotation.ExcelWriteProperty;
import com.github.excel.enums.*;
import com.github.excel.model.ExcelBaseModel;
import com.github.export.ExcelDefaultDataFormatTest;
import lombok.Data;

import java.math.BigDecimal;
import java.util.Calendar;
import java.util.Date;

/**
 * @Author: Vachel Wang
 * @Date: 2026/4/24
 * @Description: 测试demo
 */
@Data
@ExcelWrite(rowIndex = 13, themeName = ExcelThemeEnum.ZEBRA,colIndex = 0,fillType = ExcelWriterListFillTypeEnum.SHIFT,scope = ExcelWriterScopeEnum.CURRENT_SHEET,fillStyle = ExcelWriterFillStyleEnum.HORIZONTAL, nameSpace = "user"/*,titleStyleName = ExcelBasicStyle.STYLE_LIST_TITLE,contentStyleName = ExcelBasicStyle.STYLE_CONTENT*//*,mergeTitleRowNum = 1,mergeTitleColNum = 1,mergeContentColNum = 1,mergeContentRowNum = 1*/)
public class UserExcelDto extends ExcelBaseModel {
	@ExcelWriteProperty(titleName = "姓名", titleModel = ExcelWriterCellTitleModelEnum.STAND_ALONE,mergeTitleColNum = 3,mergeRowNum = 2)
	private String name;
	@ExcelWriteProperty(titleName = "性别", titleModel = ExcelWriterCellTitleModelEnum.STAND_ALONE)
	private Byte sex;
	@ExcelWriteProperty(titleName = "年龄", titleModel = ExcelWriterCellTitleModelEnum.WITH_VALUE)
	private Integer age;
	@ExcelWriteProperty(titleName = "年龄short", titleModel = ExcelWriterCellTitleModelEnum.STAND_ALONE)
	private Short ageShort;
	@ExcelWriteProperty(titleName = "身高double",colWidth=-1/*, titleStyleName = ExcelBasicStyle.STYLE_TITLE, contentStyleName = ExcelBasicStyle.STYLE_CONTENT*/, titleModel = ExcelWriterCellTitleModelEnum.WITH_VALUE, fillStyle = ExcelWriterFillStyleEnum.VERTICAL)
	private Double heightDouble;
	@ExcelWriteProperty(titleName = "身高", titleModel = ExcelWriterCellTitleModelEnum.STAND_ALONE)
	private Float height;
	@ExcelWriteProperty(titleName = "昵称",colWidth = -1, titleModel = ExcelWriterCellTitleModelEnum.STAND_ALONE,fillStyle = ExcelWriterFillStyleEnum.VERTICAL)
	private String nickName;
	@ExcelWriteProperty(titleName = "头像", titleModel = ExcelWriterCellTitleModelEnum.STAND_ALONE, linkName = "点击查看", colWidth = -1)
	private String avater;
	@ExcelWriteProperty(titleName = "邮箱", titleModel = ExcelWriterCellTitleModelEnum.STAND_ALONE,  linkName = "点击发送", colWidth = -1)
	private String email;
	@ExcelWriteProperty(titleName = "账户金额", colWidth = 200/*, titleStyleName = ExcelBasicStyle.STYLE_TITLE_RED_FONT, contentStyleName = ExcelBasicStyle.STYLE_CONTENT*/, titleModel = ExcelWriterCellTitleModelEnum.WITH_VALUE, separator = " | ")
	private Long money;
	@ExcelWriteProperty(titleName = "金额格式化", titleModel = ExcelWriterCellTitleModelEnum.STAND_ALONE, formatPattern = "###,###,###.000")
	private BigDecimal moneyBig;
	@ExcelWriteProperty(titleName = "锁定", titleModel = ExcelWriterCellTitleModelEnum.STAND_ALONE)
	private Boolean lock;
	@ExcelWriteProperty(titleName = "创建时间", formatPattern = "yyyy-MM-dd HH:mm:ss.sss", titleModel = ExcelWriterCellTitleModelEnum.STAND_ALONE, formatter = ExcelDefaultDataFormatTest.class, colWidth = -1,fillStyle = ExcelWriterFillStyleEnum.VERTICAL)
	private Date createTime;
	@ExcelWriteProperty(titleName = "修改时间", formatPattern = "yyyy-MM-dd", titleModel = ExcelWriterCellTitleModelEnum.WITH_VALUE,colWidth = 300,fillStyle = ExcelWriterFillStyleEnum.VERTICAL,verticalNewLine = true)
	private Calendar updateTime;
	@ExcelWriteProperty(titleName = "logo", titleModel = ExcelWriterCellTitleModelEnum.WITH_VALUE,/* contentStyleName = ExcelBasicStyle.STYLE_AROUND_BORDER_READ,*/ colIndex = 0, rowIndex = 0, rowHeight = 51, colWidth = 430)
	private byte[] logo;
	@ExcelWriteProperty(titleName ="公司")
	private CompanyDto company;
}
