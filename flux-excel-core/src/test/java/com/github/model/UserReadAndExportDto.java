package com.github.model;

import com.github.excel.annotation.*;
import com.github.excel.enums.ExcelThemeEnum;
import com.github.excel.exception.ExcelReaderException;
import com.github.excel.model.ExcelBaseModel;
import lombok.*;
import org.hibernate.validator.constraints.Range;

import jakarta.validation.constraints.NotNull;

/**
 * @Author: Vachel Wang
 * @Date: 2026/4/24
 * @Description: 简单导出
 */
@Data
@ExcelRead(checkTitle = true)
@ExcelWrite(themeName = ExcelThemeEnum.ZEBRA)
@ExcelValidation
@ToString
@Builder
@AllArgsConstructor
@NoArgsConstructor
public class UserReadAndExportDto extends ExcelBaseModel {
	@ExcelReadProperty(titleName = "姓名")
	@ExcelWriteProperty(titleName = "姓名")
	private String name;
	@ExcelReadProperty(titleName = "性别")
	@ExcelWriteProperty(titleName = "性别")
	private Byte sex;
	@ExcelReadProperty(titleName = "年龄")
	@ExcelWriteProperty(titleName = "年龄")
	@Range(min = 0,max = 100,message = "年龄范围只能为0-100")
	private Integer age;
	@ExcelReadProperty(titleName = "身高")
	@NotNull(message = "身高不能为空")
	@ExcelWriteProperty(titleName = "身高")
	private Float height;
	@ExcelReadProperty(titleName = "头像")
	@ExcelWriteProperty(titleName = "头像")
	private byte[] avater;
/*
	@ExcelImportProperty(titleName = "头像")
	private List<ReadPictureModel> avaterList;

	@ExcelImportProperty(titleName = "头像")
	private ReadPictureModel avaterModel;*/


	@Override
	public void callback() throws ExcelReaderException {
		// 回调方法，可进行校验、格式化等操作
		// 当解析成model时，解析完成后调用
		// 当解析成model list 时，每解析完成一个model时调用
	}

}
