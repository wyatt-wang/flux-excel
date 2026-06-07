package com.github.model;

import com.github.excel.annotation.ExcelWrite;
import com.github.excel.annotation.ExcelWriteProperty;
import com.github.excel.enums.ExcelWriterCellTitleModelEnum;
import com.github.excel.enums.ExcelWriterFillStyleEnum;
import com.github.excel.enums.ExcelWriterScopeEnum;
import com.github.excel.model.ExcelBaseModel;
import com.github.excel.write.style.ExcelBasicStyle;
import com.github.export.ExcelDefaultDataFormatTest;
import lombok.Data;

import java.math.BigDecimal;
import java.util.Calendar;
import java.util.Date;
import java.util.Map;

/**
 * @Author: Vachel Wang
 * @Date: 2026/4/24
 * @Description: 测试demo
 */
@Data
@ExcelWrite(rowIndex = 1, colIndex = 0, scope = ExcelWriterScopeEnum.CURRENT_SHEET, nameSpace = "user")
public class UserExcelDto2 extends ExcelBaseModel {
	@ExcelWriteProperty(titleName = "姓名", titleModel = ExcelWriterCellTitleModelEnum.WITH_VALUE, fillStyle = ExcelWriterFillStyleEnum.HORIZONTAL, mergeTitleColNum = 1,mergeRowNum = 1,mergeContentColNum = 1)
	private String name;
	@ExcelWriteProperty(titleName = "性别", titleModel = ExcelWriterCellTitleModelEnum.WITH_VALUE, fillStyle = ExcelWriterFillStyleEnum.HORIZONTAL,mergeTitleColNum = 2,mergeRowNum = 2,mergeContentColNum = 1)
	private Byte sex;
	@ExcelWriteProperty(titleName = "年龄", titleModel = ExcelWriterCellTitleModelEnum.WITH_VALUE, fillStyle = ExcelWriterFillStyleEnum.HORIZONTAL,mergeTitleColNum = 3,mergeRowNum = 1,mergeContentColNum = 1)
	private Integer age;
	@ExcelWriteProperty(titleName = "年龄short", titleModel = ExcelWriterCellTitleModelEnum.WITH_VALUE, fillStyle = ExcelWriterFillStyleEnum.HORIZONTAL)
	private Short ageShort;
	@ExcelWriteProperty(titleName = "身高double", titleStyleName = ExcelBasicStyle.STYLE_TITLE, contentStyleName = ExcelBasicStyle.STYLE_CONTENT, titleModel = ExcelWriterCellTitleModelEnum.STAND_ALONE, fillStyle = ExcelWriterFillStyleEnum.HORIZONTAL)
	private Double heightDouble;
	@ExcelWriteProperty(titleName = "身高", titleModel = ExcelWriterCellTitleModelEnum.WITH_VALUE, fillStyle = ExcelWriterFillStyleEnum.VERTICAL)
	private Float height;
	@ExcelWriteProperty(titleName = "昵称", titleModel = ExcelWriterCellTitleModelEnum.WITH_VALUE, fillStyle = ExcelWriterFillStyleEnum.HORIZONTAL,mergeTitleColNum = 3,mergeRowNum = 1,mergeContentColNum = 2)
	private String nickName;
	@ExcelWriteProperty(titleName = "头像", titleModel = ExcelWriterCellTitleModelEnum.WITH_VALUE, fillStyle = ExcelWriterFillStyleEnum.HORIZONTAL, linkName = "点击查看", colWidth = -1)
	private String avater;
	@ExcelWriteProperty(titleName = "邮箱", titleModel = ExcelWriterCellTitleModelEnum.STAND_ALONE, fillStyle = ExcelWriterFillStyleEnum.HORIZONTAL, linkName = "点击发送", colWidth = -1)
	private String email;
	@ExcelWriteProperty(titleName = "账户金额", colWidth = 200, titleStyleName = ExcelBasicStyle.STYLE_TITLE_RED_FONT, contentStyleName = ExcelBasicStyle.STYLE_CONTENT, titleModel = ExcelWriterCellTitleModelEnum.WITH_VALUE, separator = " | ",fillStyle = ExcelWriterFillStyleEnum.HORIZONTAL, colIndex = 20, rowIndex = 1)
	private Long money;
	@ExcelWriteProperty(titleName = "", titleModel = ExcelWriterCellTitleModelEnum.STAND_ALONE, formatPattern = "###,###,###.000", fillStyle = ExcelWriterFillStyleEnum.VERTICAL)
	private BigDecimal moneyBig;
	@ExcelWriteProperty(titleName = "锁定", titleModel = ExcelWriterCellTitleModelEnum.STAND_ALONE, fillStyle = ExcelWriterFillStyleEnum.HORIZONTAL)
	private Boolean lock;
	@ExcelWriteProperty(titleName = "创建时间", formatPattern = "yyyy-MM-dd HH:mm:ss.sss", titleModel = ExcelWriterCellTitleModelEnum.STAND_ALONE, formatter = ExcelDefaultDataFormatTest.class, fillStyle = ExcelWriterFillStyleEnum.VERTICAL, colWidth = -1, contentStyleName = ExcelBasicStyle.STYLE_AROUND_BORDER_READ)
	private Date createTime;
	@ExcelWriteProperty(titleName = "", titleModel = ExcelWriterCellTitleModelEnum.STAND_ALONE, fillStyle = ExcelWriterFillStyleEnum.HORIZONTAL, contentStyleName = ExcelBasicStyle.STYLE_AROUND_BORDER_READ, rowHeight = 51, colWidth = 430/*,mergeContentColNum = 1,mergeRowNum = 1,mergeTitleColNum = 1*/)
	private Map<String,String> map;
	@ExcelWriteProperty(titleName = "修改时间", formatPattern = "yyyy-MM-dd", titleModel = ExcelWriterCellTitleModelEnum.WITH_VALUE, fillStyle = ExcelWriterFillStyleEnum.HORIZONTAL,colWidth = 300)
	private Calendar updateTime;
	@ExcelWriteProperty(titleName = "logo", titleModel = ExcelWriterCellTitleModelEnum.WITH_VALUE, fillStyle = ExcelWriterFillStyleEnum.VERTICAL, contentStyleName = ExcelBasicStyle.STYLE_AROUND_BORDER_READ, colIndex = 0, rowIndex = 0, rowHeight = 51, colWidth = 430)
	private byte[] logo;



}
