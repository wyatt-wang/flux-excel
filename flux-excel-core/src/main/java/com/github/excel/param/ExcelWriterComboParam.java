package com.github.excel.param;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;
import lombok.experimental.Accessors;
import lombok.experimental.SuperBuilder;

import jakarta.validation.constraints.NotEmpty;
import jakarta.validation.constraints.NotNull;
import jakarta.validation.constraints.Size;
import java.util.Arrays;
import java.util.List;

/**
 * @Author: Vachel Wang
 * @Date: 2026/4/24
 * @Description: 导出验证-下拉框model
 */
@Data
@Accessors(chain = true)
@SuperBuilder
@AllArgsConstructor
@NoArgsConstructor
public final class ExcelWriterComboParam extends ExcelWriterValidationParam {
	/**
	 * 选项
	 */
	@NotNull(message = "选项不能为空")
	@Size(min = 1,message = "选项至少有一个")
	private List<String> options;

	public ExcelWriterComboParam setOptions(String[] options) {
		this.options = options == null ? null : Arrays.asList(options);
		return this;
	}

	public ExcelWriterComboParam setCommentText(String commentText) {
		setMessage(commentText);
		return this;
	}

	public ExcelWriterComboParam setCommentFontName(String commentFontName) {
		setTitle(commentFontName);
		return this;
	}
}
