package com.github.excel;

import com.github.excel.fluent.ExcelBatchWriteBuilder;
import com.github.excel.fluent.ExcelEventReadBuilder;
import com.github.excel.fluent.ExcelInspectBuilder;
import com.github.excel.fluent.ExcelLargeWriteBuilder;
import com.github.excel.fluent.ExcelReadBuilder;
import com.github.excel.fluent.ExcelWriteBuilder;

import java.io.File;
import java.io.InputStream;
import java.io.OutputStream;

/**
 * Unified fluent entry point for flux-excel v3 style APIs.
 */
public final class Excel {

	private Excel() {
	}

	public static ExcelWriteBuilder write(OutputStream outputStream) {
		return ExcelWriteBuilder.toStream(outputStream);
	}

	public static ExcelWriteBuilder write(File file) {
		return ExcelWriteBuilder.toFile(file);
	}

	public static ExcelReadBuilder read(InputStream inputStream) {
		return ExcelReadBuilder.fromStream(inputStream);
	}

	public static ExcelReadBuilder read(File file) {
		return ExcelReadBuilder.fromFile(file);
	}

	public static ExcelInspectBuilder inspect(File file) {
		return new ExcelInspectBuilder(file);
	}

	public static ExcelLargeWriteBuilder largeWrite(OutputStream outputStream) {
		return ExcelLargeWriteBuilder.toStream(outputStream);
	}

	public static ExcelLargeWriteBuilder largeWrite(File file) {
		return ExcelLargeWriteBuilder.toFile(file);
	}

	public static ExcelBatchWriteBuilder batchWrite(String outputDirPath) {
		return new ExcelBatchWriteBuilder(outputDirPath);
	}

	public static ExcelEventReadBuilder eventRead(InputStream inputStream) {
		return ExcelEventReadBuilder.fromStream(inputStream);
	}

	public static ExcelEventReadBuilder eventRead(File file) {
		return ExcelEventReadBuilder.fromFile(file);
	}
}
