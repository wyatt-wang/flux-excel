package com.github.excel.param;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;
import lombok.experimental.Accessors;
import lombok.experimental.SuperBuilder;
import org.hibernate.validator.constraints.Range;

import jakarta.validation.constraints.NotEmpty;

/**
 * @Author: Vachel Wang
 * @Date: 2026/4/24
 * @Description: 超链接model
 */
@Data
@Accessors(chain = true)
@SuperBuilder
@AllArgsConstructor
@NoArgsConstructor
public class ExcelWriterLinkParam extends ExcelWriterCellParam {
	/**
	 * 超链接名称
	 */
	private String linkName;
}
