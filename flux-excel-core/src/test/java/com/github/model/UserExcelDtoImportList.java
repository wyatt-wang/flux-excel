package com.github.model;

import com.github.excel.annotation.ExcelRead;
import com.github.excel.annotation.ExcelReadProperty;
import com.github.excel.exception.ExcelReaderException;
import com.github.excel.model.ExcelBaseModel;
import com.github.excel.model.ExcelReaderPictureModel;
import lombok.Data;

import java.util.Date;
import java.util.List;
import java.util.Objects;

/**
 * @Author: Vachel Wang
 * @Date: 2026/4/24
 * @Description: 简单导出
 */
@Data
@ExcelRead(enableSeparator = false)
public class UserExcelDtoImportList extends ExcelBaseModel {
	@ExcelReadProperty(titleName = "排序")
	private Integer sort;
	@ExcelReadProperty(titleName = "姓名")
	private String name;
	@ExcelReadProperty(titleName = "性别")
	private String sex;
	@ExcelReadProperty(titleName = "年龄")
	private String age;
	@ExcelReadProperty(titleName = "身高")
	private Float height;
	@ExcelReadProperty(titleName = "昵称")
	private String nickName;
	@ExcelReadProperty(titleName = "头像")
	private byte[] avater;
	@ExcelReadProperty(titleName = "创建时间",formatPattern = "yyyy-MM-dd")
	private Date createTime;
	@ExcelReadProperty(titleName = "备注")
	private String remark;
	@ExcelReadProperty(titleName = "logo")
	private List<ExcelReaderPictureModel> logo;

	@Override
	public String toString() {
		return "UserExcelDtoImportList{" + "name='" + name + '\'' + ", sex=" + sex + ", age=" + age + ", height=" + height + ", nickName='" + nickName + '\'' + ", avater='" + avater + '\'' + ", createTime=" + createTime + ", remark='" + remark + '\'' + '}';
	}

	@Override
	public void callback() throws ExcelReaderException {
		if (Objects.isNull(name)) {
			throw new ExcelReaderException("姓名不能为空");
		}
	}
}
