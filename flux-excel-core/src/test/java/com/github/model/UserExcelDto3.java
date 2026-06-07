package com.github.model;

import com.github.excel.annotation.ExcelWrite;
import com.github.excel.annotation.ExcelWriteProperty;
import com.github.excel.enums.ExcelWriterFillStyleEnum;
import com.github.excel.enums.ExcelThemeEnum;
import com.github.excel.param.ExcelWriterComboParam;
import com.github.excel.param.ExcelWriterCommentParam;
import com.github.excel.model.ExcelBaseModel;
import com.github.excel.param.ExcelWriterNumberScopeParam;
import com.github.excel.write.style.ExcelBasicStyle;
import lombok.Data;

import java.util.Date;

/**
 * @Author: Vachel Wang
 * @Date: 2026/4/24
 * @Description: 简单导出
 */
@Data
@ExcelWrite(themeName = ExcelThemeEnum.ZEBRA,incrementSequenceNo = true,mergeContentRowNum = 1,freezeTitle = true)
public class UserExcelDto3 extends ExcelBaseModel {
	@ExcelWriteProperty(titleName = "姓名",commentText = "",commentFontName = ExcelBasicStyle.FONT_SIZE16_BLOLD_RED)
	private String name;
	@ExcelWriteProperty(titleName = "性别",commentText = "性别只能是男女",commentFontName = ExcelBasicStyle.FONT_SIZE16_BLOLD_RED,mergeTitleColNum = 2,fillStyle = ExcelWriterFillStyleEnum.VERTICAL,colIndex = 5,rowIndex = 1)
	private Byte sex;
	@ExcelWriteProperty(titleName = "性别-1",fillStyle = ExcelWriterFillStyleEnum.VERTICAL)
	private ExcelWriterCommentParam sexStr;
	@ExcelWriteProperty(titleName = "年龄")
	private Integer age;
	@ExcelWriteProperty(titleName = "身高",fillStyle = ExcelWriterFillStyleEnum.VERTICAL,disable = true)
	private Float height;
	@ExcelWriteProperty(titleName = "昵称")
	private String nickName;
	@ExcelWriteProperty(titleName = "头像",fillStyle = ExcelWriterFillStyleEnum.VERTICAL,linkName = "点击查看")
	private String avater;
	@ExcelWriteProperty(titleName = "logo")
	private byte[] logo;
	@ExcelWriteProperty(titleName = "创建时间",fillStyle = ExcelWriterFillStyleEnum.VERTICAL,formatPattern = "yyyy年MM月dd日")
	private Date createTime;
	@ExcelWriteProperty(titleName = "公司性质",dropDownOptions = {"国企","民营企业"})
	private String companyType;
	@ExcelWriteProperty(titleName = "公司性质combo",colWidth = 100,fillStyle = ExcelWriterFillStyleEnum.VERTICAL)
	private ExcelWriterComboParam companyType1;
	@ExcelWriteProperty(titleName = "年龄范围")
	private ExcelWriterNumberScopeParam scopeModel;
}
