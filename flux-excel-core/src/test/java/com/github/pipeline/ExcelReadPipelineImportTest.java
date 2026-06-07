package com.github.pipeline;

import com.github.excel.Excel;
import com.github.excel.enums.ExcelSuffixEnum;
import com.github.excel.model.ExcelReadResult;
import com.github.excel.model.ExcelSheetInfo;
import com.github.excel.param.ExcelReaderListParam;
import com.github.excel.read.facade.ExcelReaderBatchProcess;
import com.github.excel.read.handler.reader.ExcelReader;
import com.github.excel.read.handler.reader.ExcelReaderFactory;
import com.github.model.UserReadAndExportDto;
import org.junit.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.util.ArrayList;
import java.util.List;
import java.util.concurrent.atomic.AtomicBoolean;
import java.util.concurrent.atomic.AtomicInteger;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertFalse;
import static org.junit.Assert.assertTrue;

public class ExcelReadPipelineImportTest {

	@Test
	public void readPipelineParsesListAndRunsCustomReader() {
		ByteArrayOutputStream outputStream = exportUsers(3);
		AtomicBoolean customReaderCalled = new AtomicBoolean(false);

		List<UserReadAndExportDto> parsed = Excel.read(new ByteArrayInputStream(outputStream.toByteArray()))
				.fileName("users.xlsx")
				.sheet(0)
				.list(UserReadAndExportDto.class)
				.custom(workbook -> {
					customReaderCalled.set(true);
					assertEquals("sheet", workbook.getSheetAt(0).getSheetName());
				})
				.parse()
				.getList(UserReadAndExportDto.class);

		assertTrue(customReaderCalled.get());
		assertEquals(3, parsed.size());
		assertEquals("张三1", parsed.get(0).getName());
		assertEquals(Integer.valueOf(1), parsed.get(0).getAge());
	}

	@Test
	public void factoryReaderUsesReadPipelineAndFlushesLastBatch() {
		ByteArrayOutputStream outputStream = exportUsers(3);
		AtomicInteger processedRows = new AtomicInteger(0);
		ExcelReader reader = ExcelReaderFactory.createUserReader("users.xlsx",
				new ByteArrayInputStream(outputStream.toByteArray()), null, true);

		reader.addList(ExcelReaderListParam.<UserReadAndExportDto>builder()
				.modelCla(UserReadAndExportDto.class)
				.sheetIndex(0)
				.batchProcess(new ExcelReaderBatchProcess<UserReadAndExportDto>() {
					@Override
					public int getBatchSize() {
						return 2;
					}

					@Override
					public void process(List<UserReadAndExportDto> dataList) {
						processedRows.addAndGet(dataList.size());
					}
				})
				.build());

		reader.parse();

		assertEquals(3, processedRows.get());
		assertTrue(reader.getList(UserReadAndExportDto.class).isEmpty());
	}

	@Test
	public void readModelIncludesInheritedFields() {
		ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
		Excel.write(outputStream)
				.fileName("inherit.xlsx")
				.suffix(ExcelSuffixEnum.XLSX)
				.sheet("sheet")
				.list(java.util.Collections.singletonList(new ExcelWritePipelineExportTest.InheritedExportModel("base", "child")),
						ExcelWritePipelineExportTest.InheritedExportModel.class)
				.export();

		List<ExcelWritePipelineExportTest.InheritedExportModel> parsed = Excel.read(new ByteArrayInputStream(outputStream.toByteArray()))
				.fileName("inherit.xlsx")
				.sheet(0)
				.list(ExcelWritePipelineExportTest.InheritedExportModel.class)
				.parse()
				.getList(ExcelWritePipelineExportTest.InheritedExportModel.class);

		assertEquals(1, parsed.size());
		assertEquals("base", parsed.get(0).getBaseValue());
		assertEquals("child", parsed.get(0).getChildValue());
	}

	@Test
	public void readBuilderSupportsSheetNameRangeResultListenerAndInspect() throws Exception {
		File file = File.createTempFile("flux-excel-read-options", ".xlsx");
		file.deleteOnExit();
		Excel.write(file)
				.sheet("用户")
				.list(users(3), UserReadAndExportDto.class)
				.export();

		AtomicInteger rowCount = new AtomicInteger();
		AtomicBoolean afterAll = new AtomicBoolean(false);
		ExcelReadResult<UserReadAndExportDto> result = Excel.read(file)
				.sheet("用户")
				.headRowNumber(1)
				.maxRows(1)
				.trimString(true)
				.readListener(new com.github.excel.read.listener.ExcelReadListener<UserReadAndExportDto>() {
					@Override
					public void onRow(UserReadAndExportDto row, com.github.excel.read.listener.ExcelReadListenerContext context) {
						rowCount.incrementAndGet();
					}

					@Override
					public void afterAll(com.github.excel.read.listener.ExcelReadListenerContext context) {
						afterAll.set(true);
					}
				})
				.list(UserReadAndExportDto.class)
				.parseResult(UserReadAndExportDto.class);

		List<ExcelSheetInfo> sheets = Excel.inspect(file).sheets();

		assertTrue(result.isSuccess());
		assertEquals(1, result.getData().size());
		assertEquals("张三1", result.getData().get(0).getName());
		assertEquals(1, rowCount.get());
		assertTrue(afterAll.get());
		assertEquals("用户", sheets.get(0).getName());
	}

	@Test
	public void parseResultCollectsCellErrorFromImportConverter() {
		ByteArrayOutputStream outputStream = exportUsers(1);

		ExcelReadResult<UserReadAndExportDto> result = Excel.read(new ByteArrayInputStream(outputStream.toByteArray()))
				.fileName("users.xlsx")
				.sheet("sheet")
				.fieldConverter("age", value -> {
					throw new IllegalArgumentException("年龄转换失败");
				})
				.list(UserReadAndExportDto.class)
				.parseResult(UserReadAndExportDto.class);

		assertFalse(result.isSuccess());
		assertEquals(1, result.getErrors().size());
		assertEquals("age", result.getErrors().get(0).getFieldName());
		assertEquals("年龄", result.getErrors().get(0).getTitleName());
		assertEquals(Integer.valueOf(1), result.getErrors().get(0).getRowIndex());
		assertEquals(Integer.valueOf(2), result.getErrors().get(0).getColIndex());
		assertEquals("年龄转换失败", result.getErrors().get(0).getMessage());
	}

	private ByteArrayOutputStream exportUsers(int count) {
		ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
		Excel.write(outputStream)
				.fileName("users.xlsx")
				.suffix(ExcelSuffixEnum.XLSX)
				.sheet("sheet")
				.list(users(count), UserReadAndExportDto.class)
				.export();
		return outputStream;
	}

	private List<UserReadAndExportDto> users(int count) {
		List<UserReadAndExportDto> users = new ArrayList<>();
		for (int i = 1; i <= count; i++) {
			users.add(UserReadAndExportDto.builder()
					.name("张三" + i)
					.sex((byte) 1)
					.age(i)
					.height(170F + i)
					.build());
		}
		return users;
	}
}
