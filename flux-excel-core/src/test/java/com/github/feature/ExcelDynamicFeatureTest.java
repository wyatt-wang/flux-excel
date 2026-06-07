package com.github.feature;

import com.github.excel.Excel;
import com.github.excel.annotation.ExcelRead;
import com.github.excel.annotation.ExcelReadProperty;
import com.github.excel.annotation.ExcelWrite;
import com.github.excel.annotation.ExcelWriteProperty;
import com.github.excel.enums.ExcelSuffixEnum;
import com.github.excel.model.ExcelBaseModel;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.util.Arrays;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

import static org.junit.Assert.*;

public class ExcelDynamicFeatureTest {

    @Test
    public void exportsAndReadsCsvMapList() {
        ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
        Map<String, Object> first = new LinkedHashMap<>();
        first.put("name", "Alice");
        first.put("age", 18);
        Map<String, Object> second = new LinkedHashMap<>();
        second.put("name", "Bob");
        second.put("age", 20);

        Excel.write(outputStream)
                .suffix(ExcelSuffixEnum.CSV)
                .listMap(Arrays.asList(first, second))
                .export();

        List<Map<String, Object>> rows = Excel.read(new ByteArrayInputStream(outputStream.toByteArray()))
                .fileName("users.csv")
                .mapList();

        assertEquals(2, rows.size());
        assertEquals("Alice", rows.get(0).get("name"));
        assertEquals("18", rows.get(0).get("age"));
        assertEquals("Bob", rows.get(1).get("name"));
        assertEquals("20", rows.get(1).get("age"));
    }

    @Test
    public void exportsAndReadsCsvBeanList() {
        ByteArrayOutputStream outputStream = new ByteArrayOutputStream();

        Excel.write(outputStream)
                .suffix(ExcelSuffixEnum.CSV)
                .list(Arrays.asList(new CsvDto("Alice", 18)), CsvDto.class)
                .export();

        List<CsvDto> rows = Excel.read(new ByteArrayInputStream(outputStream.toByteArray()))
                .fileName("users.csv")
                .list(CsvDto.class)
                .parse()
                .getList(CsvDto.class);

        assertEquals(1, rows.size());
        assertEquals("Alice", rows.get(0).getName());
        assertEquals(Integer.valueOf(18), rows.get(0).getAge());
    }

    @Test
    public void exportsEmptyHeaderFromMap() throws Exception {
        ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
        Map<String, String> headers = new LinkedHashMap<>();
        headers.put("name", "姓名");
        headers.put("age", "年龄");

        Excel.write(outputStream)
                .emptyHeader(headers)
                .export();

        try (Workbook workbook = WorkbookFactory.create(new ByteArrayInputStream(outputStream.toByteArray()))) {
            assertEquals("姓名", workbook.getSheetAt(0).getRow(0).getCell(0).getStringCellValue());
            assertEquals("年龄", workbook.getSheetAt(0).getRow(0).getCell(1).getStringCellValue());
        }
    }

    @Test
    public void exportsAndReadsXlsxMapList() {
        ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
        Map<String, Object> row = new LinkedHashMap<>();
        row.put("name", "Alice");
        row.put("roles", Arrays.asList("admin", "user"));

        Excel.write(outputStream)
                .listMap(Arrays.asList(row))
                .export();

        List<Map<String, Object>> rows = Excel.read(new ByteArrayInputStream(outputStream.toByteArray()))
                .fileName("users.xlsx")
                .mapList();

        assertEquals(1, rows.size());
        assertEquals("Alice", rows.get(0).get("name"));
        assertEquals("[admin, user]", rows.get(0).get("roles"));
    }

    @Test
    public void exportsEmptyHeaderFromBeanClass() throws Exception {
        ByteArrayOutputStream outputStream = new ByteArrayOutputStream();

        Excel.write(outputStream)
                .emptyHeader(SequenceDto.class)
                .export();

        try (Workbook workbook = WorkbookFactory.create(new ByteArrayInputStream(outputStream.toByteArray()))) {
            assertEquals("姓名", workbook.getSheetAt(0).getRow(0).getCell(0).getStringCellValue());
        }
    }

    @Test
    public void exportsBeanListWithRuntimeAliasIncludeAndConverter() throws Exception {
        ByteArrayOutputStream outputStream = new ByteArrayOutputStream();

        Excel.write(outputStream)
                .table(CsvDto.class)
                .alias("name", "用户名称")
                .include("name")
                .fieldConverter("name", value -> "用户-" + value)
                .list(Arrays.asList(new CsvDto("Alice", 18)))
                .export();

        try (Workbook workbook = WorkbookFactory.create(new ByteArrayInputStream(outputStream.toByteArray()))) {
            assertEquals("用户名称", workbook.getSheetAt(0).getRow(0).getCell(0).getStringCellValue());
            assertEquals("用户-Alice", workbook.getSheetAt(0).getRow(1).getCell(0).getStringCellValue());
            assertNull(workbook.getSheetAt(0).getRow(0).getCell(1));
        }
    }

    @Test
    public void standardWriterAppliesRuntimeAliasIncludeAndConverter() throws Exception {
        ByteArrayOutputStream outputStream = new ByteArrayOutputStream();

        Excel.write(outputStream)
                .alias("age", "年龄段")
                .onlyAlias(true)
                .fieldConverter("age", value -> "A" + value)
                .list(Arrays.asList(new CsvDto("Alice", 18)), CsvDto.class)
                .export();

        try (Workbook workbook = WorkbookFactory.create(new ByteArrayInputStream(outputStream.toByteArray()))) {
            assertEquals("年龄段", workbook.getSheetAt(0).getRow(0).getCell(0).getStringCellValue());
            assertEquals("A18", workbook.getSheetAt(0).getRow(1).getCell(0).getStringCellValue());
            assertNull(workbook.getSheetAt(0).getRow(0).getCell(1));
        }
    }

    @Test
    public void appliesTextWatermarkToXlsx() throws Exception {
        ByteArrayOutputStream outputStream = new ByteArrayOutputStream();

        Excel.write(outputStream)
                .watermark("DRAFT")
                .list(Arrays.asList(new SequenceDto("A")), SequenceDto.class)
                .export();

        try (XSSFWorkbook workbook = new XSSFWorkbook(new ByteArrayInputStream(outputStream.toByteArray()))) {
            XSSFSheet sheet = workbook.getSheetAt(0);
            assertTrue(sheet.getCTWorksheet().isSetPicture());
        }
    }

    @Test
    public void sequenceNumberRestartsForEachTableBlockInSameSheet() throws Exception {
        ByteArrayOutputStream outputStream = new ByteArrayOutputStream();

        Excel.write(outputStream)
                .list(Arrays.asList(new SequenceDto("A"), new SequenceDto("B")), SequenceDto.class)
                .at(5, 0)
                .list(Arrays.asList(new SequenceDto("C"), new SequenceDto("D")), SequenceDto.class)
                .export();

        try (Workbook workbook = WorkbookFactory.create(new ByteArrayInputStream(outputStream.toByteArray()))) {
            assertEquals(1d, workbook.getSheetAt(0).getRow(1).getCell(0).getNumericCellValue(), 0d);
            assertEquals(2d, workbook.getSheetAt(0).getRow(2).getCell(0).getNumericCellValue(), 0d);
            assertEquals(1d, workbook.getSheetAt(0).getRow(6).getCell(0).getNumericCellValue(), 0d);
            assertEquals(2d, workbook.getSheetAt(0).getRow(7).getCell(0).getNumericCellValue(), 0d);
        }
    }

    @ExcelWrite(incrementSequenceNo = true)
    public static class SequenceDto extends ExcelBaseModel {
        @ExcelWriteProperty(titleName = "姓名")
        private String name;

        public SequenceDto() {
        }

        public SequenceDto(String name) {
            this.name = name;
        }

        public String getName() {
            return name;
        }

        public void setName(String name) {
            this.name = name;
        }
    }

    @ExcelWrite
    @ExcelRead
    public static class CsvDto extends ExcelBaseModel {
        @ExcelWriteProperty(titleName = "姓名")
        @ExcelReadProperty(titleName = "姓名")
        private String name;

        @ExcelWriteProperty(titleName = "年龄")
        @ExcelReadProperty(titleName = "年龄")
        private Integer age;

        public CsvDto() {
        }

        public CsvDto(String name, Integer age) {
            this.name = name;
            this.age = age;
        }

        public String getName() {
            return name;
        }

        public void setName(String name) {
            this.name = name;
        }

        public Integer getAge() {
            return age;
        }

        public void setAge(Integer age) {
            this.age = age;
        }
    }
}
