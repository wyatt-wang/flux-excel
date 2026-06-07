package com.github.excel.fluent;

import com.github.excel.model.ExcelSheetInfo;

import java.io.File;
import java.util.List;
import java.util.Map;

/**
 * Excel inspection entry for sheet metadata and small data previews.
 */
public class ExcelInspectBuilder {

	private final File file;
	private String fileName;

	public ExcelInspectBuilder(File file) {
		this.file = file;
		this.fileName = file == null ? null : file.getName();
	}

	public ExcelInspectBuilder fileName(String fileName) {
		this.fileName = fileName;
		return this;
	}

	public List<ExcelSheetInfo> sheets() {
		return ExcelReadBuilder.fromFile(file)
				.fileName(fileName)
				.sheets();
	}

	public List<Map<String, Object>> preview(int rows) {
		return ExcelReadBuilder.fromFile(file)
				.fileName(fileName)
				.preview(rows);
	}
}
