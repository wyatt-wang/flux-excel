package com.github.pipeline;

import com.github.excel.Excel;
import com.github.excel.annotation.ExcelRead;
import com.github.excel.annotation.ExcelReadProperty;
import com.github.excel.annotation.ExcelWrite;
import com.github.excel.annotation.ExcelWriteProperty;
import com.github.excel.enums.ExcelSuffixEnum;
import com.github.excel.model.ExcelBaseModel;
import com.github.excel.param.ExcelWriterConditionalStyleParam;
import com.github.excel.param.ExcelWriterCellParam;
import com.github.excel.param.ExcelWriterMergeParam;
import org.apache.poi.ss.usermodel.ConditionalFormatting;
import org.apache.poi.ss.usermodel.FillPatternType;
import com.github.excel.write.ExcelWriter;
import org.apache.poi.ss.usermodel.Hyperlink;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.SheetConditionalFormatting;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Rule;
import org.junit.Test;
import org.junit.rules.TemporaryFolder;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.util.Collections;
import java.util.zip.ZipEntry;
import java.util.zip.ZipInputStream;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertNotNull;

public class ExcelWritePipelineExportTest {

	@Rule
	public TemporaryFolder temporaryFolder = new TemporaryFolder();

	@Test
	public void pipelineWritesCustomCellsAndPostProcessor() throws Exception {
		ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
		ExcelWriterCellParam cellParam = new ExcelWriterCellParam();
		cellParam.setValue("marker");
		cellParam.setSheetName("sheet");
		cellParam.setRowIndex(0);
		cellParam.setColIndex(0);

		Excel.write(outputStream)
				.column(cellParam)
				.custom(workbook -> workbook.getSheetAt(0).createRow(1).createCell(0).setCellValue("post"))
				.export();

		try (Workbook workbook = WorkbookFactory.create(new ByteArrayInputStream(outputStream.toByteArray()))) {
			assertEquals("marker", workbook.getSheetAt(0).getRow(0).getCell(0).getStringCellValue());
			assertEquals("post", workbook.getSheetAt(0).getRow(1).getCell(0).getStringCellValue());
		}
	}

	@Test
	public void pipelineWritesMergedCustomCells() throws Exception {
		ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
		ExcelWriterMergeParam mergeParam = new ExcelWriterMergeParam()
				.setEndRowIndex(0)
				.setEndColIndex(1);
		mergeParam.setValue("merged")
				.setSheetName("sheet")
				.setRowIndex(0)
				.setColIndex(0);

		Excel.write(outputStream)
				.mergeColumn(mergeParam)
				.export();

		try (Workbook workbook = WorkbookFactory.create(new ByteArrayInputStream(outputStream.toByteArray()))) {
			assertEquals("merged", workbook.getSheetAt(0).getRow(0).getCell(0).getStringCellValue());
			assertEquals(1, workbook.getSheetAt(0).getNumMergedRegions());
		}
	}

	@Test
	public void writerClearsCollectedStateAfterPipelineExport() throws Exception {
		ByteArrayOutputStream firstOutputStream = new ByteArrayOutputStream();
		ExcelWriter writer = Excel.write(firstOutputStream)
				.noneDataTips(false)
				.column(cell("first"))
				.writer();

		writer.export();

		ByteArrayOutputStream secondOutputStream = new ByteArrayOutputStream();
		writer.export(secondOutputStream, "second.xlsx", com.github.excel.enums.ExcelSuffixEnum.XLSX);

		try (Workbook workbook = WorkbookFactory.create(new ByteArrayInputStream(secondOutputStream.toByteArray()))) {
			Row firstRow = workbook.getSheetAt(0).getRow(0);
			assertEquals("", firstRow.getCell(0).getStringCellValue());
		}
	}

	@Test
	public void writePropertySupportsFriendlySizeAndDynamicHyperlinkName() throws Exception {
		ByteArrayOutputStream outputStream = new ByteArrayOutputStream();

		Excel.write(outputStream)
				.fileName("links.xlsx")
				.suffix(ExcelSuffixEnum.XLSX)
				.list(Collections.singletonList(new HyperlinkExportModel("https://example.com", "Example")), HyperlinkExportModel.class)
				.export();

		try (Workbook workbook = WorkbookFactory.create(new ByteArrayInputStream(outputStream.toByteArray()))) {
			Row row = workbook.getSheetAt(0).getRow(1);
			assertEquals(24F, row.getHeightInPoints(), 0.1F);
			assertEquals("Example", row.getCell(0).getStringCellValue());
			Hyperlink hyperlink = row.getCell(0).getHyperlink();
			assertNotNull(hyperlink);
			assertEquals("https://example.com", hyperlink.getAddress());
			assertEquals(20 * 256, workbook.getSheetAt(0).getColumnWidth(0), 128);
		}
	}

	@Test
	public void exportModelIncludesInheritedFields() throws Exception {
		ByteArrayOutputStream outputStream = new ByteArrayOutputStream();

		Excel.write(outputStream)
				.fileName("inherit.xlsx")
				.suffix(ExcelSuffixEnum.XLSX)
				.list(Collections.singletonList(new InheritedExportModel("base", "child")), InheritedExportModel.class)
				.export();

		try (Workbook workbook = WorkbookFactory.create(new ByteArrayInputStream(outputStream.toByteArray()))) {
			assertEquals("父字段", workbook.getSheetAt(0).getRow(0).getCell(0).getStringCellValue());
			assertEquals("子字段", workbook.getSheetAt(0).getRow(0).getCell(1).getStringCellValue());
			assertEquals("base", workbook.getSheetAt(0).getRow(1).getCell(0).getStringCellValue());
			assertEquals("child", workbook.getSheetAt(0).getRow(1).getCell(1).getStringCellValue());
		}
	}

	@Test
	public void waterfallWriteAppendsListsWithoutCoordinates() throws Exception {
		ByteArrayOutputStream outputStream = new ByteArrayOutputStream();

		Excel.write(outputStream)
				.waterfall()
				.list(Collections.singletonList(new InheritedExportModel("base-1", "child-1")), InheritedExportModel.class)
				.list(Collections.singletonList(new InheritedExportModel("base-2", "child-2")), InheritedExportModel.class)
				.export();

		try (Workbook workbook = WorkbookFactory.create(new ByteArrayInputStream(outputStream.toByteArray()))) {
			assertEquals("父字段", workbook.getSheetAt(0).getRow(0).getCell(0).getStringCellValue());
			assertEquals("base-1", workbook.getSheetAt(0).getRow(1).getCell(0).getStringCellValue());
			assertEquals("父字段", workbook.getSheetAt(0).getRow(2).getCell(0).getStringCellValue());
			assertEquals("base-2", workbook.getSheetAt(0).getRow(3).getCell(0).getStringCellValue());
		}
	}

	@Test
	public void writeListSupportsFieldConditionalStyle() throws Exception {
		ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
		ExcelWriterConditionalStyleParam conditionalStyleParam = ExcelWriterConditionalStyleParam.greaterThan("10")
				.fill(IndexedColors.YELLOW)
				.font(IndexedColors.RED)
				.setBold(true);

		Excel.write(outputStream)
				.list(Collections.singletonList(new ConditionalStyleExportModel("item", 12)), ConditionalStyleExportModel.class)
				.conditionalStyle("amount", conditionalStyleParam)
				.export();

		try (Workbook workbook = WorkbookFactory.create(new ByteArrayInputStream(outputStream.toByteArray()))) {
			SheetConditionalFormatting formatting = workbook.getSheetAt(0).getSheetConditionalFormatting();
			assertEquals(1, formatting.getNumConditionalFormattings());
			ConditionalFormatting conditionalFormatting = formatting.getConditionalFormattingAt(0);
			CellRangeAddress[] ranges = conditionalFormatting.getFormattingRanges();
			assertEquals(1, ranges.length);
			assertEquals(1, ranges[0].getFirstRow());
			assertEquals(1, ranges[0].getLastRow());
			assertEquals(1, ranges[0].getFirstColumn());
			assertEquals(1, ranges[0].getLastColumn());
		}
	}

	@Test
	public void writerAppliesTemplateFillAndLightStyleDsl() throws Exception {
		ByteArrayOutputStream outputStream = new ByteArrayOutputStream();

		Excel.write(outputStream)
				.custom(workbook -> workbook.getSheetAt(0).createRow(3).createCell(0).setCellValue("{title}"))
				.fill("title", "报表标题")
				.headerStyle((cellStyle, workbook) -> {
					cellStyle.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
					cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
				})
				.columnStyle("金额", (cellStyle, workbook) -> {
					cellStyle.setFillForegroundColor(IndexedColors.LIGHT_GREEN.getIndex());
					cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
				})
				.list(Collections.singletonList(new ConditionalStyleExportModel("item", 12)), ConditionalStyleExportModel.class)
				.export();

		try (Workbook workbook = WorkbookFactory.create(new ByteArrayInputStream(outputStream.toByteArray()))) {
			assertEquals("报表标题", workbook.getSheetAt(0).getRow(3).getCell(0).getStringCellValue());
			assertEquals(IndexedColors.YELLOW.getIndex(), workbook.getSheetAt(0).getRow(0).getCell(0).getCellStyle().getFillForegroundColor());
			assertEquals(IndexedColors.LIGHT_GREEN.getIndex(), workbook.getSheetAt(0).getRow(1).getCell(1).getCellStyle().getFillForegroundColor());
		}
	}

	@Test
	public void largeWriteBuilderExportsThroughPipeline() throws Exception {
		ByteArrayOutputStream outputStream = new ByteArrayOutputStream();

		Excel.largeWrite(outputStream)
				.sheet("large")
				.list(Collections.singletonList(new InheritedExportModel("base", "child")), InheritedExportModel.class)
				.export();

		try (Workbook workbook = WorkbookFactory.create(new ByteArrayInputStream(outputStream.toByteArray()))) {
			assertEquals("large-1", workbook.getSheetAt(0).getSheetName());
			assertEquals("父字段", workbook.getSheetAt(0).getRow(0).getCell(0).getStringCellValue());
			assertEquals("base", workbook.getSheetAt(0).getRow(1).getCell(0).getStringCellValue());
		}
	}

	@Test
	public void streamingWriteUsesConfigurableSxssfOptions() throws Exception {
		ByteArrayOutputStream outputStream = new ByteArrayOutputStream();

		Excel.write(outputStream)
				.streaming(true)
				.rowAccessWindowSize(1)
				.compressTempFiles(true)
				.useSharedStringsTable(true)
				.list(Collections.singletonList(new InheritedExportModel("base", "child")), InheritedExportModel.class)
				.export();

		try (Workbook workbook = WorkbookFactory.create(new ByteArrayInputStream(outputStream.toByteArray()))) {
			assertEquals("父字段", workbook.getSheetAt(0).getRow(0).getCell(0).getStringCellValue());
			assertEquals("base", workbook.getSheetAt(0).getRow(1).getCell(0).getStringCellValue());
		}
	}

	@Test
	public void streamingTemplateWriteWrapsXlsxTemplate() throws Exception {
		File template = temporaryFolder.newFile("template.xlsx");
		try (Workbook workbook = new XSSFWorkbook()) {
			workbook.createSheet("sheet").createRow(0).createCell(0).setCellValue("template");
			try (java.io.FileOutputStream outputStream = new java.io.FileOutputStream(template)) {
				workbook.write(outputStream);
			}
		}

		ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
		Excel.write(outputStream)
				.templateFilePath(template.getAbsolutePath())
				.streaming(true)
				.rowAccessWindowSize(1)
				.compressTempFiles(true)
				.useSharedStringsTable(true)
				.custom(workbook -> workbook.getSheet("sheet").createRow(1).createCell(0).setCellValue("streaming-template"))
				.export();

		try (Workbook workbook = WorkbookFactory.create(new ByteArrayInputStream(outputStream.toByteArray()))) {
			assertEquals("template", workbook.getSheet("sheet").getRow(0).getCell(0).getStringCellValue());
			assertEquals("streaming-template", workbook.getSheet("sheet").getRow(1).getCell(0).getStringCellValue());
		}
	}

	@Test
	public void largeWriteBuilderAcceptsStreamingOptions() throws Exception {
		ByteArrayOutputStream outputStream = new ByteArrayOutputStream();

		Excel.largeWrite(outputStream)
				.sheet("large")
				.rowAccessWindowSize(1)
				.compressTempFiles(true)
				.useSharedStringsTable(true)
				.list(Collections.singletonList(new InheritedExportModel("base", "child")), InheritedExportModel.class)
				.export();

		try (Workbook workbook = WorkbookFactory.create(new ByteArrayInputStream(outputStream.toByteArray()))) {
			assertEquals("large-1", workbook.getSheetAt(0).getSheetName());
			assertEquals("base", workbook.getSheetAt(0).getRow(1).getCell(0).getStringCellValue());
		}
	}

	@Test
	public void batchWriteBuilderExportsZipThroughPipeline() throws Exception {
		File outputDir = temporaryFolder.newFolder("batch");

		Excel.batchWrite(outputDir.getAbsolutePath())
				.file("batch-data")
				.sheet("batch")
				.list(Collections.singletonList(new InheritedExportModel("base", "child")), InheritedExportModel.class)
				.exportZip("batch-export");

		File zipFile = new File(outputDir, "batch-export.zip");
		try (ZipInputStream zipInputStream = new ZipInputStream(new FileInputStream(zipFile))) {
			ZipEntry entry = zipInputStream.getNextEntry();
			assertNotNull(entry);
			assertEquals("batch-data.xlsx", entry.getName());
			try (Workbook workbook = WorkbookFactory.create(zipInputStream)) {
				assertEquals("batch", workbook.getSheetAt(0).getSheetName());
				assertEquals("父字段", workbook.getSheetAt(0).getRow(0).getCell(0).getStringCellValue());
				assertEquals("base", workbook.getSheetAt(0).getRow(1).getCell(0).getStringCellValue());
			}
		}
	}

	@Test
	public void batchWriteBuilderAcceptsStreamingOptions() throws Exception {
		File outputDir = temporaryFolder.newFolder("batch-options");

		Excel.batchWrite(outputDir.getAbsolutePath())
				.rowAccessWindowSize(1)
				.compressTempFiles(true)
				.useSharedStringsTable(true)
				.file("batch-options-data")
				.sheet("batch")
				.list(Collections.singletonList(new InheritedExportModel("base", "child")), InheritedExportModel.class)
				.exportZip("batch-options-export");

		File zipFile = new File(outputDir, "batch-options-export.zip");
		try (ZipInputStream zipInputStream = new ZipInputStream(new FileInputStream(zipFile))) {
			ZipEntry entry = zipInputStream.getNextEntry();
			assertNotNull(entry);
			assertEquals("batch-options-data.xlsx", entry.getName());
			try (Workbook workbook = WorkbookFactory.create(zipInputStream)) {
				assertEquals("batch", workbook.getSheetAt(0).getSheetName());
				assertEquals("base", workbook.getSheetAt(0).getRow(1).getCell(0).getStringCellValue());
			}
		}
	}

	private ExcelWriterCellParam cell(String value) {
		ExcelWriterCellParam cellParam = new ExcelWriterCellParam();
		cellParam.setValue(value);
		cellParam.setSheetName("sheet");
		cellParam.setRowIndex(0);
		cellParam.setColIndex(0);
		return cellParam;
	}

	public static class BaseInheritedExportModel extends ExcelBaseModel {
		@ExcelReadProperty(titleName = "父字段")
		@ExcelWriteProperty(titleName = "父字段", index = 1)
		private String baseValue;

		public String getBaseValue() {
			return baseValue;
		}

		public void setBaseValue(String baseValue) {
			this.baseValue = baseValue;
		}
	}

	@ExcelRead
	@ExcelWrite
	public static class InheritedExportModel extends BaseInheritedExportModel {
		@ExcelReadProperty(titleName = "子字段")
		@ExcelWriteProperty(titleName = "子字段", index = 2)
		private String childValue;

		public InheritedExportModel() {
		}

		public InheritedExportModel(String baseValue, String childValue) {
			setBaseValue(baseValue);
			this.childValue = childValue;
		}

		public String getChildValue() {
			return childValue;
		}

		public void setChildValue(String childValue) {
			this.childValue = childValue;
		}
	}

	@ExcelWrite
	public static class HyperlinkExportModel extends ExcelBaseModel {
		@ExcelWriteProperty(titleName = "链接", rowHeightPoints = 24, colWidthChars = 20, linkNameField = "displayName")
		private String url;

		private String displayName;

		public HyperlinkExportModel() {
		}

		public HyperlinkExportModel(String url, String displayName) {
			this.url = url;
			this.displayName = displayName;
		}

		public String getUrl() {
			return url;
		}

		public void setUrl(String url) {
			this.url = url;
		}

		public String getDisplayName() {
			return displayName;
		}

		public void setDisplayName(String displayName) {
			this.displayName = displayName;
		}
	}

	@ExcelWrite
	public static class ConditionalStyleExportModel extends ExcelBaseModel {
		@ExcelWriteProperty(titleName = "名称", index = 1)
		private String name;

		@ExcelWriteProperty(titleName = "金额", index = 2)
		private Integer amount;

		public ConditionalStyleExportModel() {
		}

		public ConditionalStyleExportModel(String name, Integer amount) {
			this.name = name;
			this.amount = amount;
		}

		public String getName() {
			return name;
		}

		public void setName(String name) {
			this.name = name;
		}

		public Integer getAmount() {
			return amount;
		}

		public void setAmount(Integer amount) {
			this.amount = amount;
		}
	}
}
