package com.github.excel.write.pipeline;

import com.github.excel.model.ExcelBaseModel;
import com.github.excel.model.ExcelWriterModel;
import com.github.excel.param.ExcelWriterCellParam;
import com.github.excel.param.ExcelWriterCommentParam;
import com.github.excel.param.ExcelWriterConditionalStyleParam;
import com.github.excel.param.ExcelWriterMergeParam;
import com.github.excel.param.ExcelWriterParam;
import com.github.excel.write.ExcelCustomWriter;
import com.github.excel.write.style.AbstractExcelStyle;
import com.github.excel.pipeline.core.CleanupAwarePipelineContext;
import lombok.Builder;
import lombok.Data;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import java.io.OutputStream;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.function.Function;

@Data
@Builder
public class ExcelWriteContext implements CleanupAwarePipelineContext {

	private OutputStream outputStream;
	private String excelName;
	private ExcelWriterParam writerParam;
	private boolean streaming;
	private boolean noneDataTips;
	private String selectSheet;
	private Class<? extends ExcelBaseModel> listCla;
	private int sheetRowMaxCount;
	private ExcelCustomWriter customWrite;
	private List<ExcelWriterModel> exportModelList;
	private List<ExcelWriterModel> exportBeanList;
	private List<ExcelWriterCellParam> customColumnModelList;
	private List<ExcelWriterMergeParam> mergeCustomColumnModelList;
	private List<Class<? extends AbstractExcelStyle>> styleList;
	private Map<Class<? extends ExcelBaseModel>, Map<String, String>> excludeFieldMap;
	private Map<Class<? extends ExcelBaseModel>, Map<String, ExcelWriterCommentParam>> commentMap;
	private Map<Class<? extends ExcelBaseModel>, Map<String, List<ExcelWriterConditionalStyleParam>>> conditionalStyleMap;
	private List<ExcelWriterConditionalStyleParam> conditionalStyleList;
	private Map<String, String> runtimeHeaderAliases;
	private Set<String> runtimeIncludeFields;
	private boolean runtimeOnlyAlias;
	private Map<Class<?>, Function<Object, Object>> runtimeTypeConverters;
	private Map<String, Function<Object, Object>> runtimeFieldConverters;
	private Workbook workbook;
	private SXSSFWorkbook sxssfWorkbook;
	private CreationHelper creationHelper;
	private boolean cleanupExecuted;
}
