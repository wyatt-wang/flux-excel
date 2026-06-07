package com.github.excel.write;

import com.github.excel.model.ExcelBaseModel;
import com.github.excel.enums.ExcelSuffixEnum;
import com.github.excel.param.ExcelWriterCellParam;
import com.github.excel.param.ExcelWriterCommentParam;
import com.github.excel.param.ExcelWriterComboParam;
import com.github.excel.param.ExcelWriterConditionalStyleParam;
import com.github.excel.param.ExcelWriterMergeParam;
import com.github.excel.param.ExcelWriterListParam;
import com.github.excel.param.ExcelWriterModelParam;
import com.github.excel.param.ExcelWriterNumberScopeParam;
import com.github.excel.write.style.AbstractExcelStyle;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;

import jakarta.servlet.http.HttpServletRequest;
import jakarta.servlet.http.HttpServletResponse;
import java.io.File;
import java.io.OutputStream;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.function.Function;

/**
 * @Author: Vachel Wang
 * @Date: 2026/4/24
 * @Description: 导出接口
 */
public interface ExcelWriter {
	/**
	 * 写入单个model
	 * @param param 参数
	 * @param <T>
	 */
	<T extends ExcelBaseModel> ExcelWriter writeModel(ExcelWriterModelParam<T> param);

	/**
	 * 写入list
	 * @param param
	 * @param <T>
	 */
	<T extends ExcelBaseModel> ExcelWriter writeList(ExcelWriterListParam<T> param);

	/**
	 * 添加自定义单元格
	 * @param customColumnModel
	 */
	ExcelWriter writeColumn(ExcelWriterCellParam customColumnModel);

	/**
	 * 合并单元格并设值
	 * @param mergeCustomColumnModel
	 */
	ExcelWriter writeMergeColumn(ExcelWriterMergeParam mergeCustomColumnModel);

	/**
	 * excel 自定义写
	 * @param customWrite
	 */
	ExcelWriter writeCustom(ExcelCustomWriter customWrite);

	/**
	 * 添加样式
	 * @param styles 样式
	 */
	<T extends AbstractExcelStyle> ExcelWriter addStyles(Class<T> ... styles);

	/**
	 * 获取样式
	 * @param name 样式名称
	 * @return
	 */
	CellStyle getStyle(String name);

	/**
	 * 获取字体
	 * @param name 字体名称
	 * @return
	 */
	Font getFont(String name);

	/**
	 * 根据 writerParam 输出
	 */
	ExcelWriter export();

	/**
	 * 导出到流
	 */
	ExcelWriter export(OutputStream outputStream, String fileName, ExcelSuffixEnum suffixEnum);

	/**
	 * 导出到响应
	 */
	ExcelWriter export(HttpServletRequest request, HttpServletResponse response, String fileName, ExcelSuffixEnum suffixEnum);

	/**
	 * 向后兼容旧 API
	 */
	ExcelWriter process(OutputStream outputStream, String fileName, ExcelSuffixEnum suffixEnum);

	/**
	 * 向后兼容旧 API
	 */
	ExcelWriter process(HttpServletRequest request, HttpServletResponse response, String fileName, ExcelSuffixEnum suffixEnum);

	/**
	 * 向后兼容旧 API
	 */
	ExcelWriter process(File file, String fileName, ExcelSuffixEnum suffixEnum);

	/**
	 * 兼容旧 API
	 */
	default <T extends ExcelBaseModel> ExcelWriter addModel(T model, String sheetName, boolean fillTemplate, Class<T> modelCla) {
		return writeModel(ExcelWriterModelParam.<T>builder()
				.model(model)
				.modelCla(modelCla)
				.sheetName(sheetName)
				.fillTemplate(fillTemplate)
				.build());
	}

	/**
	 * 兼容旧 API
	 */
	default <T extends ExcelBaseModel> ExcelWriter addModel(int rowIndex, int colIndex, T model, String sheetName, boolean fillTemplate, Class<T> modelCla) {
		return writeModel(ExcelWriterModelParam.<T>builder()
				.model(model)
				.modelCla(modelCla)
				.sheetName(sheetName)
				.fillTemplate(fillTemplate)
				.rowIndex(rowIndex)
				.colIndex(colIndex)
				.build());
	}

	/**
	 * 兼容旧 API
	 */
	default <T extends ExcelBaseModel> ExcelWriter addModelList(List<T> modelList, String sheetName, boolean fillTemplate, Class<T> modelCla) {
		return writeList(ExcelWriterListParam.<T>builder()
				.modelList(modelList)
				.modelCla(modelCla)
				.sheetName(sheetName)
				.fillTemplate(fillTemplate)
				.build());
	}

	/**
	 * 兼容旧 API
	 */
	default <T extends ExcelBaseModel> ExcelWriter addModelList(int rowIndex, int colIndex, List<T> modelList, String sheetName, boolean fillTemplate, Class<T> modelCla) {
		return writeList(ExcelWriterListParam.<T>builder()
				.modelList(modelList)
				.modelCla(modelCla)
				.sheetName(sheetName)
				.fillTemplate(fillTemplate)
				.rowIndex(rowIndex)
				.colIndex(colIndex)
				.build());
	}

	/**
	 * 兼容旧 API
	 */
	ExcelWriter excludes(String[] excludeFields);

	/**
	 * 兼容旧 API
	 */
	ExcelWriter addValidationOrComment(String field, ExcelWriterCommentParam commentParam);

	/**
	 * 兼容旧 API
	 */
	ExcelWriter addValidationOrComment(String field, ExcelWriterNumberScopeParam validationParam);

	/**
	 * 兼容旧 API
	 */
	ExcelWriter addValidationOrComment(String field, ExcelWriterComboParam validationParam);

	/**
	 * 兼容旧 API
	 */
	ExcelWriter addConditionalStyle(String field, ExcelWriterConditionalStyleParam conditionalStyleParam);

	/**
	 * 兼容旧 API
	 */
	ExcelWriter addConditionalStyle(ExcelWriterConditionalStyleParam conditionalStyleParam);

	/**
	 * 兼容旧 API
	 */
	ExcelWriter setNoneDataTips(boolean noneDataTips);

	/**
	 * 兼容旧 API
	 */
	ExcelWriter setStreaming(boolean streaming);

	/**
	 * 兼容旧 API
	 */
	default ExcelWriter setCustomWrite(ExcelCustomWriter customWrite) {
		return writeCustom(customWrite);
	}

	/**
	 * 兼容旧 API
	 */
	ExcelWriter selectSheet(String sheetName);

	/**
	 * 兼容旧 API
	 */
	ExcelWriter setListCla(Class<? extends ExcelBaseModel> listCla);

	/**
	 * 运行时表头别名
	 */
	ExcelWriter runtimeHeaderAliases(Map<String, String> aliases);

	/**
	 * 运行时 include 字段
	 */
	ExcelWriter runtimeIncludeFields(Set<String> fields);

	/**
	 * 仅导出配置了别名的字段
	 */
	ExcelWriter runtimeOnlyAlias(boolean onlyAlias);

	/**
	 * 运行时类型转换器
	 */
	ExcelWriter runtimeTypeConverters(Map<Class<?>, Function<Object, Object>> converters);

	/**
	 * 运行时字段转换器
	 */
	ExcelWriter runtimeFieldConverters(Map<String, Function<Object, Object>> converters);

	/**
	 * 兼容旧 API
	 */
	default ExcelWriter addCustomColumn(ExcelWriterCellParam customColumnModel) {
		return writeColumn(customColumnModel);
	}

	/**
	 * 兼容旧 API
	 */
	default ExcelWriter addMergeCustomColumn(ExcelWriterMergeParam mergeCustomColumnModel) {
		return writeMergeColumn(mergeCustomColumnModel);
	}

	/**
	 * 兼容旧 API
	 */
	default <T extends AbstractExcelStyle> ExcelWriter addStyle(Class<T> style) {
		return addStyles(style);
	}

}
