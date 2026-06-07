package com.github.excel.read.pipeline;

import com.github.excel.context.ExcelReaderContext;
import com.github.excel.context.ExcelReaderModelContext;
import com.github.excel.model.ExcelBaseModel;
import com.github.excel.pipeline.core.CleanupAwarePipelineContext;
import lombok.Builder;
import lombok.Data;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Workbook;

import java.util.List;
import java.util.Map;

@Data
@Builder
public class ExcelReadContext<T extends ExcelBaseModel> implements CleanupAwarePipelineContext {

	private ExcelReaderContext<T> readerContext;
	private Workbook workbook;
	private FormulaEvaluator formulaEvaluator;
	private Map<Integer, List<ExcelReaderModelContext<T>>> sheetModelMap;
	private boolean cleanupExecuted;
}
