package com.github.model;

import com.github.excel.annotation.ExcelRead;
import com.github.excel.annotation.ExcelReadProperty;
import com.github.excel.model.ExcelBaseModel;
import com.github.excel.model.ExcelReaderPictureModel;
import lombok.Data;

import java.util.Date;
import java.util.List;

/**
 * @Author: Vachel Wang
 * @Date: 2026/4/24
 * @Description: 简单导出
 */
@Data
@ExcelRead(enableSeparator = false)
public class UserExcelDtoImportList1 extends ExcelBaseModel {
	@ExcelReadProperty(titleName = "姓名")
	private String name;
	@ExcelReadProperty(titleName = "性别")
	private String sex;
	@ExcelReadProperty(titleName = "年龄")
	private Integer age;
	@ExcelReadProperty(titleName = "身高")
	private Float height;
	@ExcelReadProperty(titleName = "昵称")
	private String nickName;
	@ExcelReadProperty(titleName = "头像")
	private byte[] avater;
	@ExcelReadProperty(titleName = "创建时间")
	private Date createTime;
	@ExcelReadProperty(titleName = "备注")
	private String remark;
	@ExcelReadProperty(titleName = "国籍")
	private String contry;
	@ExcelReadProperty(titleName = "logo")
	private List<ExcelReaderPictureModel> logo;

	@Override
	public String toString() {
		return "UserExcelDtoImportList1{" + "name='" + name + '\'' + ", sex='" + sex + '\'' + ", age=" + age + ", height=" + height + ", nickName='" + nickName + '\'' + ", avater='" + avater + '\'' + ", createTime=" + createTime + ", remark='" + remark + '\'' + ", contry='" + contry + '\'' + '}';
	}
}
