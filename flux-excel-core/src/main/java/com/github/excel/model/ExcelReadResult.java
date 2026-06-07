package com.github.excel.model;

import lombok.Data;
import lombok.experimental.Accessors;

import java.util.ArrayList;
import java.util.List;

/**
 * Excel 导入结构化结果
 */
@Data
@Accessors(chain = true)
public class ExcelReadResult<T extends ExcelBaseModel> {
	private List<T> data = new ArrayList<>();
	private List<ExcelReadError> errors = new ArrayList<>();
	private int totalRows;
	private int validRows;
	private int invalidRows;

	public boolean isSuccess() {
		return errors == null || errors.isEmpty();
	}
}
