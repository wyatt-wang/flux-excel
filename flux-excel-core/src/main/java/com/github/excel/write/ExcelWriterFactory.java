package com.github.excel.write;

import com.github.excel.exception.ExcelWriterException;
import com.github.excel.model.ExcelBaseModel;
import com.github.excel.enums.ExcelSuffixEnum;
import com.github.excel.param.ExcelWriterParam;
import com.github.excel.param.ExcelWriterSteamParam;
import com.github.excel.util.StringUtil;
import com.github.excel.write.impl.ExcelLargeListBatchWriterImpl;
import com.github.excel.write.impl.ExcelLargeListWriterImpl;
import com.github.excel.write.impl.ExcelUserWriterImpl;

/**
 * @Author: Vachel Wang
 * @Date: 2026/4/24
 * @Description: excel导出工厂
 */
public class ExcelWriterFactory {
	private static final String DEFAULT_TEMPLATE = null;

	/**
	 * 创建用户模式导出
	 *
	 * @return ExcelWriter
	 */
	public static ExcelWriter createUserModelWriter() {
		return createUserModelWriter(DEFAULT_TEMPLATE);
	}

	/**
	 * 创建用户模式导出
	 *
	 * @param template 模版名称
	 * @return ExcelWriter
	 */
	public static ExcelWriter createUserModelWriterWithTemplate(String template) {
		return createUserModelWriter(template);
	}

	private static ExcelWriter createUserModelWriter(String template) {
		ExcelWriterSteamParam writerParam = ExcelWriterSteamParam.builder()
				.outputStream(new java.io.ByteArrayOutputStream())
				.suffixEnum(resolveSuffix(template))
				.template(template)
				.build();
		return createUserModelWriter(writerParam);
	}

	/**
	 * 创建用户模式导出
	 *
	 * @return ExcelWriter
	 */
	public static ExcelWriter createUserModelWriter(ExcelWriterParam writerParam) {
		return new ExcelUserWriterImpl(writerParam);
	}

	private static ExcelSuffixEnum resolveSuffix(String template) {
		if (StringUtil.notEmpty(template) && template.toLowerCase().endsWith(ExcelSuffixEnum.XLS.getSuffix())) {
			return ExcelSuffixEnum.XLS;
		}
		return ExcelSuffixEnum.XLSX;
	}

	/**
	 * 创建大数据模式导出
	 *
	 * @param sheetName sheet名称
	 * @return ExcelWriter
	 */
	public static ExcelLargeListWriter createLargeListWriter(String sheetName) {
		return new ExcelLargeListWriterImpl(sheetName);
	}

	public static ExcelLargeListWriter createLargeListWriter(String sheetName, int rowAccessWindowSize,
															boolean compressTempFiles, boolean useSharedStringsTable) {
		return new ExcelLargeListWriterImpl(sheetName, rowAccessWindowSize, compressTempFiles, useSharedStringsTable);
	}

	/**
	 * 创建大数据导出并设置sheet最大行数
	 *
	 * @param sheetName
	 * @param sheetRowMaxCount
	 * @return
	 */
	public static ExcelLargeListWriter createLargeListWriter(String sheetName, int sheetRowMaxCount) {
		return new ExcelLargeListWriterImpl(sheetName, sheetRowMaxCount);
	}

	public static ExcelLargeListWriter createLargeListWriter(String sheetName, int sheetRowMaxCount,
															int rowAccessWindowSize, boolean compressTempFiles,
															boolean useSharedStringsTable) {
		return new ExcelLargeListWriterImpl(sheetName, sheetRowMaxCount, rowAccessWindowSize, compressTempFiles,
				useSharedStringsTable);
	}

	/**
	 * 创建大数据导出并设置sheet最大行数
	 *
	 * @param sheetName sheet 名称
	 * @param sheetRowMaxCount sheet 最大行数
	 * @param listCla 列表class
	 * @return
	 */
	public static ExcelLargeListWriter createLargeListWriter(String sheetName, int sheetRowMaxCount,Class<? extends ExcelBaseModel> listCla) {
		return new ExcelLargeListWriterImpl(sheetName, sheetRowMaxCount,listCla);
	}

	public static ExcelLargeListWriter createLargeListWriter(String sheetName, int sheetRowMaxCount,
															Class<? extends ExcelBaseModel> listCla,
															int rowAccessWindowSize, boolean compressTempFiles,
															boolean useSharedStringsTable) {
		return new ExcelLargeListWriterImpl(sheetName, sheetRowMaxCount, listCla, rowAccessWindowSize,
				compressTempFiles, useSharedStringsTable);
	}

	/**
	 * 创建大数据模式导出
	 *
	 * @param outputDirPath 导出文件夹名
	 * @return ExcelWriter
	 */
	public static ExcelLargeListBatchWriter createLargeListBatchWriter(String outputDirPath) {
		return new ExcelLargeListBatchWriterImpl(outputDirPath);
	}

	public static ExcelLargeListBatchWriter createLargeListBatchWriter(String outputDirPath, int rowAccessWindowSize,
																	  boolean compressTempFiles,
																	  boolean useSharedStringsTable) {
		return new ExcelLargeListBatchWriterImpl(outputDirPath, null, rowAccessWindowSize, compressTempFiles,
				useSharedStringsTable);
	}

	/**
	 * 创建大数据模式导出
	 *
	 * @param outputDirPath 导出文件夹名
	 * @param maxPoolSize   线程池大小
	 * @return ExcelWriter
	 */
	public static ExcelLargeListBatchWriter createLargeListBatchWriter(String outputDirPath, int maxPoolSize) {
		return new ExcelLargeListBatchWriterImpl(outputDirPath, maxPoolSize);
	}

	public static ExcelLargeListBatchWriter createLargeListBatchWriter(String outputDirPath, int maxPoolSize,
																	  int rowAccessWindowSize,
																	  boolean compressTempFiles,
																	  boolean useSharedStringsTable) {
		return new ExcelLargeListBatchWriterImpl(outputDirPath, maxPoolSize, rowAccessWindowSize, compressTempFiles,
				useSharedStringsTable);
	}

}
