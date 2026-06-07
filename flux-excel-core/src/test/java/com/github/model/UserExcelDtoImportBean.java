package com.github.model;

import com.github.excel.annotation.ExcelValidation;
import com.github.excel.annotation.ExcelRead;
import com.github.excel.annotation.ExcelReadProperty;
import com.github.excel.exception.ExcelReaderException;
import com.github.excel.model.ExcelBaseModel;
import com.github.read.ValidateGroup;
import lombok.Data;
import lombok.ToString;
import lombok.extern.slf4j.Slf4j;
import org.hibernate.validator.constraints.Range;

import jakarta.validation.constraints.NotEmpty;
import jakarta.validation.constraints.NotNull;
import jakarta.validation.groups.Default;

/**
 * @Author: Vachel Wang
 * @Date: 2026/4/24
 * @Description: 简单导出
 */
@Data
@ExcelRead(checkTitle = true)
@ExcelValidation
@Slf4j
@ToString
public class UserExcelDtoImportBean extends ExcelBaseModel {
	@ExcelReadProperty(titleName = "姓名")
	@NotEmpty(message = "名称不能为空",groups = {ValidateGroup.NameGroup.class})
	private String name;
	@ExcelReadProperty(titleName = "性别")
	@NotNull(message = "性别不能为空")
	private Byte sex;
	@ExcelReadProperty(titleName = "年龄")
	@NotNull(message = "年龄不能为空",groups = {ValidateGroup.AgeGroup.class})
	@Range(min = 0,max = 100,message = "年龄范围只能为0-100",groups = {ValidateGroup.AgeGroup.class})
	private String age;
	@ExcelReadProperty(titleName = "身高")
	@NotNull(message = "身高不能为空")
	private Float height;
	/*@ExcelImportProperty(titleName = "昵称")
	private String nickName;
	@ExcelImportProperty(titleName = "头像")
	private String avater;
	@ExcelImportProperty(titleName = "logo")
	private ReadPictureModel logo;
	@ExcelImportProperty(titleName = "创建时间",formatPattern = "yyyy-MM-dd")
	private Date createTime;
	@ExcelImportProperty(titleName = "备注")
	private String remark;
	@ExcelImportProperty(titleName = "余额")
	private BigDecimal money;
	@ExcelImportProperty(titleName = "余额BigInteger")
	private BigInteger moneyInt;
	@ExcelImportProperty(titleName = "小孩")
	private Boolean hasChlid;

	@Override
	public String toString() {
		return "UserExcelDtoImportBean{" + "name='" + name + '\'' + ", sex=" + sex + ", age=" + age + ", height=" + height + ", nickName='" + nickName + '\'' + ", avater='" + avater + '\'' + ", logo=" + logo + ", createTime=" + createTime + ", remark='" + remark + '\'' + ", money=" + money + ", moneyInt=" + moneyInt + '}';
	}*/

	@Override
	public void callback() throws ExcelReaderException {
		log.info(this.toString());
	}

	@Override
	public Class<?>[] validationGroup() {
		if(sex.equals(new Integer(1).byteValue())) {
			return new Class[]{ValidateGroup.NameGroup.class, Default.class};
		}else {
			return new Class[]{ValidateGroup.AgeGroup.class, Default.class};
		}
	}
}
