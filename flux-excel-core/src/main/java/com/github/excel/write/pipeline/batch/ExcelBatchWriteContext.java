package com.github.excel.write.pipeline.batch;

import com.github.excel.model.ExcelBaseModel;
import com.github.excel.write.ExcelLargeListBatchWriter;
import com.github.excel.write.style.AbstractExcelStyle;
import lombok.Data;

import java.util.ArrayList;
import java.util.List;

@Data
public class ExcelBatchWriteContext {

	private ExcelLargeListBatchWriter writer;
	private String zipFileName;
	private List<Class<? extends AbstractExcelStyle>> styles = new ArrayList<>();
	private List<ListOperation> operations = new ArrayList<>();

	@Data
	public static class ListOperation {
		private List<? extends ExcelBaseModel> modelList;
		private String fileName;
		private String sheetName;
		private Class<? extends ExcelBaseModel> modelClass;
		private String[] excludeFields;
	}
}
