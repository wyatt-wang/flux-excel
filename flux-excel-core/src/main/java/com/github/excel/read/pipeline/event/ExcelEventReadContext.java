package com.github.excel.read.pipeline.event;

import com.github.excel.read.facade.AbstractEventBatchHandler;
import com.github.excel.read.handler.event.ExcelEventReader;
import com.github.excel.read.handler.event.ExcelEventRowReader;
import lombok.Data;

import java.io.File;
import java.io.InputStream;

@Data
public class ExcelEventReadContext<T> {

	private InputStream inputStream;
	private File file;
	private String fileName;
	private ExcelEventRowReader<T> rowReader;
	private AbstractEventBatchHandler<T> batchHandler;
	private ExcelEventReader<T> eventReader;
}
