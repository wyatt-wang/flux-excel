package com.github.excel.util;

import com.github.excel.enums.ExcelSuffixEnum;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
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
		return readWorkbook(inputStream);
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

	private static List<Map<String, Object>> readWorkbook(InputStream inputStream) {
		try (Workbook workbook = WorkbookFactory.create(inputStream)) {
			Sheet sheet = workbook.getSheetAt(0);
			if (sheet == null || sheet.getPhysicalNumberOfRows() == 0) {
				return new ArrayList<>();
			}
			DataFormatter formatter = new DataFormatter();
			Row headerRow = sheet.getRow(0);
			List<String> headers = new ArrayList<>();
			for (int i = 0; headerRow != null && i < headerRow.getLastCellNum(); i++) {
				headers.add(formatter.formatCellValue(headerRow.getCell(i)));
			}
			List<Map<String, Object>> rows = new ArrayList<>();
			for (int i = 1; i <= sheet.getLastRowNum(); i++) {
				Row row = sheet.getRow(i);
				if (row == null) {
					continue;
				}
				Map<String, Object> rowData = new LinkedHashMap<>();
				for (int j = 0; j < headers.size(); j++) {
					Cell cell = row.getCell(j);
					rowData.put(headers.get(j), formatter.formatCellValue(cell));
				}
				rows.add(rowData);
			}
			return rows;
		} catch (Exception e) {
			throw new IllegalStateException("Read excel map list failed", e);
		}
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
		try (Reader reader = new InputStreamReader(inputStream, StandardCharsets.UTF_8);
			 BufferedReader bufferedReader = new BufferedReader(reader)) {
			List<List<String>> records = parseCsv(bufferedReader);
			if (records.isEmpty()) {
				return new ArrayList<>();
			}
			List<String> headers = records.get(0);
			List<Map<String, Object>> rows = new ArrayList<>();
			for (int i = 1; i < records.size(); i++) {
				List<String> record = records.get(i);
				Map<String, Object> row = new LinkedHashMap<>();
				for (int j = 0; j < headers.size(); j++) {
					row.put(headers.get(j), j < record.size() ? record.get(j) : "");
				}
				rows.add(row);
			}
			return rows;
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

	private static List<List<String>> parseCsv(BufferedReader reader) throws IOException {
		List<List<String>> records = new ArrayList<>();
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
				records.add(record);
				record = new ArrayList<>();
			} else {
				field.append(ch);
			}
		}
		if (field.length() > 0 || !record.isEmpty()) {
			record.add(trimCarriageReturn(field));
			records.add(record);
		}
		return records;
	}

	private static String trimCarriageReturn(StringBuilder field) {
		int length = field.length();
		if (length > 0 && field.charAt(length - 1) == '\r') {
			return field.substring(0, length - 1);
		}
		return field.toString();
	}
}
