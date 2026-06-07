package com.github.excel.util;

import com.github.excel.enums.ExcelSuffixEnum;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.BufferedReader;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.io.OutputStream;
import java.io.OutputStreamWriter;
import java.io.Reader;
import java.io.UncheckedIOException;
import java.io.Writer;
import java.nio.charset.StandardCharsets;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.Set;
import java.util.concurrent.atomic.AtomicReference;
import java.util.function.Consumer;

public final class DynamicExcelUtil {

	private DynamicExcelUtil() {
	}

	public static void writeMapList(OutputStream outputStream, ExcelSuffixEnum suffix, String sheetName,
									List<? extends Map<String, ?>> rows, Map<String, String> headers,
									String watermarkText) {
		if (ExcelSuffixEnum.CSV == suffix) {
			writeCsv(outputStream, rows, headers);
			return;
		}
		writeWorkbook(outputStream, suffix, sheetName, rows, headers, watermarkText);
	}

	public static List<Map<String, Object>> readMapList(InputStream inputStream, String fileName) {
		if (fileName != null && fileName.toLowerCase().endsWith(ExcelSuffixEnum.CSV.getSuffix())) {
			return readCsv(inputStream);
		}
		return readWorkbook(inputStream, 0, null, 0, 1, null, null, null, false, true, true);
	}

	public static List<Map<String, Object>> readMapList(InputStream inputStream, String fileName, int sheetIndex,
														String sheetName, Integer headRowNumber, Integer dataStartRow,
														Integer dataEndRow, Integer maxRows, boolean trimString,
														boolean ignoreEmptyRow, boolean fillMergedCells) {
		return readMapList(inputStream, fileName, sheetIndex, sheetName, headRowNumber, 1, dataStartRow,
				dataEndRow, maxRows, trimString, ignoreEmptyRow, fillMergedCells);
	}

	public static List<Map<String, Object>> readMapList(InputStream inputStream, String fileName, int sheetIndex,
														String sheetName, Integer headRowNumber, Integer headRowCount,
														Integer dataStartRow, Integer dataEndRow, Integer maxRows,
														boolean trimString, boolean ignoreEmptyRow,
														boolean fillMergedCells) {
		if (fileName != null && fileName.toLowerCase().endsWith(ExcelSuffixEnum.CSV.getSuffix())) {
			List<Map<String, Object>> rows = readCsv(inputStream);
			return sliceRows(rows, dataStartRow, dataEndRow, maxRows);
		}
		return readWorkbook(inputStream, sheetIndex, sheetName, headRowNumber, headRowCount, dataStartRow, dataEndRow, maxRows,
				trimString, ignoreEmptyRow, fillMergedCells);
	}

	private static void writeWorkbook(OutputStream outputStream, ExcelSuffixEnum suffix, String sheetName,
									  List<? extends Map<String, ?>> rows, Map<String, String> headers,
									  String watermarkText) {
		try (Workbook workbook = ExcelSuffixEnum.XLS == suffix ? new HSSFWorkbook() : new XSSFWorkbook()) {
			Sheet sheet = workbook.createSheet(sheetName == null || sheetName.trim().isEmpty() ? "sheet" : sheetName);
			List<String> keys = resolveKeys(rows, headers);
			writeHeader(sheet, keys, headers);
			writeRows(sheet, rows, keys);
			ExcelWatermarkUtil.applyTextWatermark(workbook, watermarkText);
			workbook.write(outputStream);
		} catch (IOException e) {
			throw new UncheckedIOException(e);
		}
	}

	private static void writeHeader(Sheet sheet, List<String> keys, Map<String, String> headers) {
		if (keys.isEmpty()) {
			return;
		}
		Row row = sheet.createRow(0);
		for (int i = 0; i < keys.size(); i++) {
			String key = keys.get(i);
			row.createCell(i).setCellValue(headers != null && headers.containsKey(key) ? headers.get(key) : key);
		}
	}

	private static void writeRows(Sheet sheet, List<? extends Map<String, ?>> rows, List<String> keys) {
		if (rows == null || rows.isEmpty()) {
			return;
		}
		for (int rowIndex = 0; rowIndex < rows.size(); rowIndex++) {
			Row row = sheet.createRow(rowIndex + 1);
			Map<String, ?> rowData = rows.get(rowIndex);
			for (int colIndex = 0; colIndex < keys.size(); colIndex++) {
				Object value = rowData.get(keys.get(colIndex));
				if (Objects.nonNull(value)) {
					row.createCell(colIndex).setCellValue(String.valueOf(value));
				}
			}
		}
	}

	private static List<Map<String, Object>> readWorkbook(InputStream inputStream, int sheetIndex, String sheetName,
														  Integer headRowNumber, Integer headRowCount, Integer dataStartRow,
														  Integer dataEndRow, Integer maxRows,
														  boolean trimString, boolean ignoreEmptyRow,
														  boolean fillMergedCells) {
		try (Workbook workbook = WorkbookFactory.create(inputStream)) {
			Sheet sheet = sheetName == null || sheetName.trim().isEmpty() ? workbook.getSheetAt(sheetIndex) : workbook.getSheet(sheetName);
			if (sheet == null || sheet.getPhysicalNumberOfRows() == 0) {
				return new ArrayList<>();
			}
			DataFormatter formatter = new DataFormatter();
			Map<Long, Cell> mergedFirstCellIndex = fillMergedCells ? buildMergedFirstCellIndex(sheet) : null;
			int headerRowIndex = headRowNumber == null ? sheet.getFirstRowNum() : headRowNumber;
			int headerRows = headRowCount == null || headRowCount < 1 ? 1 : headRowCount;
			int firstDataRow = dataStartRow == null ? headerRowIndex + headerRows : dataStartRow;
			int lastDataRow = dataEndRow == null ? sheet.getLastRowNum() : Math.min(dataEndRow, sheet.getLastRowNum());
			Row headerRow = sheet.getRow(headerRowIndex + headerRows - 1);
			List<String> headers = new ArrayList<>();
			for (int i = 0; headerRow != null && i < headerRow.getLastCellNum(); i++) {
				headers.add(resolveHeader(sheet, headerRowIndex, headerRows, i, formatter, mergedFirstCellIndex, trimString));
			}
			List<Map<String, Object>> rows = new ArrayList<>();
			for (int i = firstDataRow; i <= lastDataRow; i++) {
				if (maxRows != null && rows.size() >= maxRows) {
					break;
				}
				Row row = sheet.getRow(i);
				if (row == null) {
					continue;
				}
				Map<String, Object> rowData = new LinkedHashMap<>();
				boolean emptyRow = true;
				for (int j = 0; j < headers.size(); j++) {
					String value = trimString(formatCell(sheet, i, j, formatter, mergedFirstCellIndex), trimString);
					if (value != null && !value.isEmpty()) {
						emptyRow = false;
					}
					rowData.put(headers.get(j), value);
				}
				if (!emptyRow || !ignoreEmptyRow) {
					rows.add(rowData);
				}
			}
			return rows;
		} catch (Exception e) {
			throw new IllegalStateException("Read excel map list failed", e);
		}
	}

	private static String resolveHeader(Sheet sheet, int headerRowIndex, int headerRows, int colIndex,
										DataFormatter formatter, Map<Long, Cell> mergedFirstCellIndex,
										boolean trimString) {
		List<String> names = new ArrayList<>();
		for (int rowIndex = headerRowIndex; rowIndex < headerRowIndex + headerRows; rowIndex++) {
			String header = trimString(formatCell(sheet, rowIndex, colIndex, formatter, mergedFirstCellIndex), trimString);
			if (header == null || header.isEmpty()) {
				continue;
			}
			if (names.isEmpty() || !header.equals(names.get(names.size() - 1))) {
				names.add(header);
			}
		}
		return String.join(".", names);
	}

	private static String formatCell(Sheet sheet, int rowIndex, int colIndex, DataFormatter formatter,
									 Map<Long, Cell> mergedFirstCellIndex) {
		Row row = sheet.getRow(rowIndex);
		Cell cell = row == null ? null : row.getCell(colIndex);
		String value = formatter.formatCellValue(cell);
		if (mergedFirstCellIndex == null || (value != null && !value.isEmpty())) {
			return value;
		}
		Cell mergedFirstCell = mergedFirstCellIndex.get(cellKey(rowIndex, colIndex));
		return mergedFirstCell == null ? value : formatter.formatCellValue(mergedFirstCell);
	}

	private static Map<Long, Cell> buildMergedFirstCellIndex(Sheet sheet) {
		Map<Long, Cell> index = new LinkedHashMap<>();
		for (int i = 0; i < sheet.getNumMergedRegions(); i++) {
			CellRangeAddress range = sheet.getMergedRegion(i);
			Row firstRow = sheet.getRow(range.getFirstRow());
			Cell firstCell = firstRow == null ? null : firstRow.getCell(range.getFirstColumn());
			if (firstCell == null) {
				continue;
			}
			for (int rowIndex = range.getFirstRow(); rowIndex <= range.getLastRow(); rowIndex++) {
				for (int colIndex = range.getFirstColumn(); colIndex <= range.getLastColumn(); colIndex++) {
					index.put(cellKey(rowIndex, colIndex), firstCell);
				}
			}
		}
		return index;
	}

	private static long cellKey(int rowIndex, int colIndex) {
		return (((long) rowIndex) << 32) | (colIndex & 0xffffffffL);
	}

	private static String trimString(String value, boolean trimString) {
		return trimString && value != null ? value.trim() : value;
	}

	private static List<Map<String, Object>> sliceRows(List<Map<String, Object>> rows, Integer dataStartRow,
													   Integer dataEndRow, Integer maxRows) {
		if (rows == null || rows.isEmpty()) {
			return rows;
		}
		int start = dataStartRow == null ? 0 : Math.max(0, dataStartRow - 1);
		int end = dataEndRow == null ? rows.size() : Math.min(rows.size(), dataEndRow);
		if (start >= end) {
			return new ArrayList<>();
		}
		List<Map<String, Object>> sliced = new ArrayList<>(rows.subList(start, end));
		if (maxRows != null && sliced.size() > maxRows) {
			return new ArrayList<>(sliced.subList(0, maxRows));
		}
		return sliced;
	}

	private static void writeCsv(OutputStream outputStream, List<? extends Map<String, ?>> rows,
								 Map<String, String> headers) {
		try (Writer writer = new OutputStreamWriter(outputStream, StandardCharsets.UTF_8)) {
			List<String> keys = resolveKeys(rows, headers);
			writeCsvLine(writer, resolveTitles(keys, headers));
			if (rows != null) {
				for (Map<String, ?> row : rows) {
					List<String> values = new ArrayList<>();
					for (String key : keys) {
						Object value = row.get(key);
						values.add(value == null ? "" : String.valueOf(value));
					}
					writeCsvLine(writer, values);
				}
			}
		} catch (IOException e) {
			throw new UncheckedIOException(e);
		}
	}

	private static List<Map<String, Object>> readCsv(InputStream inputStream) {
		List<Map<String, Object>> rows = new ArrayList<>();
		readCsv(inputStream, rows::add);
		return rows;
	}

	public static void readCsv(InputStream inputStream, Consumer<Map<String, Object>> rowConsumer) {
		try (Reader reader = new InputStreamReader(inputStream, StandardCharsets.UTF_8);
			 BufferedReader bufferedReader = new BufferedReader(reader)) {
			AtomicReference<List<String>> headers = new AtomicReference<>();
			forEachCsvRecord(bufferedReader, record -> {
				if (headers.get() == null) {
					headers.set(record);
					return;
				}
				Map<String, Object> row = new LinkedHashMap<>();
				for (int j = 0; j < headers.get().size(); j++) {
					row.put(headers.get().get(j), j < record.size() ? record.get(j) : "");
				}
				rowConsumer.accept(row);
			});
		} catch (IOException e) {
			throw new UncheckedIOException(e);
		}
	}

	private static List<String> resolveKeys(List<? extends Map<String, ?>> rows, Map<String, String> headers) {
		Set<String> keys = new LinkedHashSet<>();
		if (headers != null) {
			keys.addAll(headers.keySet());
		}
		if (rows != null) {
			for (Map<String, ?> row : rows) {
				if (row != null) {
					keys.addAll(row.keySet());
				}
			}
		}
		return new ArrayList<>(keys);
	}

	private static List<String> resolveTitles(List<String> keys, Map<String, String> headers) {
		List<String> titles = new ArrayList<>();
		for (String key : keys) {
			titles.add(headers != null && headers.containsKey(key) ? headers.get(key) : key);
		}
		return titles;
	}

	private static void writeCsvLine(Writer writer, List<String> values) throws IOException {
		for (int i = 0; i < values.size(); i++) {
			if (i > 0) {
				writer.write(',');
			}
			writer.write(escapeCsv(values.get(i)));
		}
		writer.write("\r\n");
	}

	private static String escapeCsv(String value) {
		if (value == null) {
			return "";
		}
		boolean quote = value.indexOf(',') >= 0 || value.indexOf('"') >= 0 || value.indexOf('\n') >= 0 || value.indexOf('\r') >= 0;
		if (!quote) {
			return value;
		}
		return "\"" + value.replace("\"", "\"\"") + "\"";
	}

	private static void forEachCsvRecord(BufferedReader reader, Consumer<List<String>> recordConsumer) throws IOException {
		List<String> record = new ArrayList<>();
		StringBuilder field = new StringBuilder();
		boolean inQuotes = false;
		int value;
		while ((value = reader.read()) != -1) {
			char ch = (char) value;
			if (inQuotes) {
				if (ch == '"') {
					reader.mark(1);
					int next = reader.read();
					if (next == '"') {
						field.append('"');
					} else {
						inQuotes = false;
						if (next != -1) {
							reader.reset();
						}
					}
				} else {
					field.append(ch);
				}
			} else if (ch == '"') {
				inQuotes = true;
			} else if (ch == ',') {
				record.add(field.toString());
				field.setLength(0);
			} else if (ch == '\n') {
				record.add(trimCarriageReturn(field));
				field.setLength(0);
				recordConsumer.accept(record);
				record = new ArrayList<>();
			} else {
				field.append(ch);
			}
		}
		if (field.length() > 0 || !record.isEmpty()) {
			record.add(trimCarriageReturn(field));
			recordConsumer.accept(record);
		}
	}

	private static String trimCarriageReturn(StringBuilder field) {
		int length = field.length();
		if (length > 0 && field.charAt(length - 1) == '\r') {
			return field.substring(0, length - 1);
		}
		return field.toString();
	}
}
