package com.github.excel.read.pipeline.execution;

import lombok.Builder;
import lombok.Data;

@Data
@Builder
public class ExcelReadExecutionContext {

	private boolean csvFile;
	private boolean csvListRead;
	private Runnable csvParser;
	private Runnable workbookParser;
	private Runnable rowNotifier;
	private Runnable afterAllNotifier;
}
