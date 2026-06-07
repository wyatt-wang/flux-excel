package com.github.excel.param;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;
import lombok.experimental.Accessors;
import lombok.experimental.SuperBuilder;

import jakarta.validation.constraints.NotEmpty;
import jakarta.validation.constraints.NotNull;

/**
 * @Author: Vachel Wang
 * @Date: 2026/4/24
 * @Description: 导出验证-数字范围
 */
@Data
@Accessors(chain = true)
@SuperBuilder
@AllArgsConstructor
@NoArgsConstructor
public  class ExcelWriterNumberScopeParam extends ExcelWriterValidationParam {
	/**
	 * 开始值
	 */
	@NotNull(message = "开始值不能为空")
	private String start;
	/**
	 * 结束值
	 */
	@NotNull(message = "结束值不能为空")
	private String end;

	public ExcelWriterNumberScopeParam setCommentText(String commentText) {
		setMessage(commentText);
		return this;
	}
}
