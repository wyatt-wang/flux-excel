package com.github.model;

import com.github.excel.annotation.ExcelRead;
import com.github.excel.annotation.ExcelReadProperty;
import com.github.excel.model.ExcelBaseModel;
import lombok.Data;

import java.util.Date;

/**
 * @Author: Vachel Wang
 * @Date: 2026/4/24
 * @Description: 简单导出
 */
@Data
@ExcelRead
public class UserExcelDtoImportBean3 extends ExcelBaseModel {
	@ExcelReadProperty(titleName = "姓名",checkNull = true)
	private String name;
	@ExcelReadProperty(titleName = "性别",checkNull = true)
	private Byte sex;
	@ExcelReadProperty(titleName = "年龄",checkNull = true)
	private Integer age;
	@ExcelReadProperty(titleName = "身高")
	private Float height;
	@ExcelReadProperty(titleName = "昵称",checkNull = true)
	private String nickName;
	@ExcelReadProperty(titleName = "头像")
	private String avater;
	@ExcelReadProperty(titleName = "创建时间",formatPattern = "yyyy-MM-dd")
	private Date createTime;
	@ExcelReadProperty(titleName = "备注",checkNull = true)
	private String remark;

	@Override
	public String toString() {
		return "UserExcelDtoImportBean2{" + "name='" + name + '\'' + ", sex=" + sex + ", age=" + age + ", height=" + height + ", nickName='" + nickName + '\'' + ", avater='" + avater + '\'' + ", createTime=" + createTime + ", remark='" + remark + '\'' + '}';
	}
}
