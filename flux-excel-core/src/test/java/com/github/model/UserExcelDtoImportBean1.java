package com.github.model;

import com.github.excel.annotation.ExcelValidation;
import com.github.excel.annotation.ExcelRead;
import com.github.excel.annotation.ExcelReadProperty;
import com.github.excel.model.ExcelBaseModel;
import lombok.Data;

import jakarta.validation.constraints.NotNull;

/**
 * @Author: Vachel Wang
 * @Date: 2026/4/24
 * @Description: 简单导出
 */
@Data
@ExcelRead(checkTitle = true)
@ExcelValidation
public class UserExcelDtoImportBean1 extends ExcelBaseModel {
	@ExcelReadProperty(titleName = "姓名")
	private String name;
	@ExcelReadProperty(titleName = "性别",checkNull = true)
	private Byte sex;
	@ExcelReadProperty(titleName = "年龄")
	private Integer age;
	@ExcelReadProperty(titleName = "身高")
	@NotNull(message = "身高不能为空")
	private Float height;
	/*@ExcelImportProperty(titleName = "昵称",checkNull = true)
	private String nickName;
	@ExcelImportProperty(titleName = "头像")
	private String avater;
	*//*@ExcelImportProperty(titleName = "logo")
	private byte[] logo;*//*
	@ExcelImportProperty(titleName = "创建时间",formatPattern = "yyyy-MM-dd")
	private Date createTime;
	@ExcelImportProperty(titleName = "备注",checkNull = true)
	private String remark;
	@ExcelImportProperty(titleName = "国籍",checkNull = true)
	private String contry;

	@Override
	public String toString() {
		return "UserExcelDtoImportBean1{" + "name='" + name + '\'' + ", sex=" + sex + ", age=" + age + ", height=" + height + ", nickName='" + nickName + '\'' + ", avater='" + avater + '\'' + ", createTime=" + createTime + ", remark='" + remark + '\'' + ", contry='" + contry + '\'' + '}';
	}*/
}
