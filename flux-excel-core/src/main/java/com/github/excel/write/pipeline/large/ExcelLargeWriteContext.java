package com.github.excel.write.pipeline.large;

import com.github.excel.model.ExcelBaseModel;
import com.github.excel.write.ExcelLargeListWriter;
import com.github.excel.write.style.AbstractExcelStyle;
import lombok.Data;

import java.io.File;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.List;

@Data
public class ExcelLargeWriteContext {

	private ExcelLargeListWriter writer;
	private OutputStream outputStream;
	private File file;
	private List<Class<? extends AbstractExcelStyle>> styles = new ArrayList<>();
	private List<ListOperation> operations = new ArrayList<>();

	@Data
	public static class ListOperation {
		private List<? extends ExcelBaseModel> modelList;
		private Class<? extends ExcelBaseModel> modelClass;
		private String[] excludeFields;
	}
}
