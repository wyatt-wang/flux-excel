package com.github.feature;

import com.github.excel.Excel;
import com.github.excel.annotation.ExcelRead;
import com.github.excel.annotation.ExcelReadProperty;
import com.github.excel.annotation.ExcelWrite;
import com.github.excel.annotation.ExcelWriteProperty;
import com.github.excel.enums.ExcelSuffixEnum;
import com.github.excel.model.ExcelBaseModel;
import com.github.excel.model.ExcelCellComment;
import com.github.excel.model.ExcelCellHyperlink;
import com.github.excel.model.ExcelHeaderInfo;
import com.github.excel.model.ExcelMergedCell;
import com.github.excel.param.ExcelRichTextValue;
import org.apache.poi.common.usermodel.HyperlinkType;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.Comment;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.Hyperlink;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellRangeAddress;
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

    @Test
    public void exportsAndReadsComplexHeaders() throws Exception {
        ByteArrayOutputStream outputStream = new ByteArrayOutputStream();

        Excel.write(outputStream)
                .list(Arrays.asList(new ComplexHeaderDto("Alice", 18, "OK")), ComplexHeaderDto.class)
                .export();

        try (Workbook workbook = WorkbookFactory.create(new ByteArrayInputStream(outputStream.toByteArray()))) {
            assertEquals("基础信息", workbook.getSheetAt(0).getRow(0).getCell(0).getStringCellValue());
            assertEquals("姓名", workbook.getSheetAt(0).getRow(1).getCell(0).getStringCellValue());
            assertEquals("年龄", workbook.getSheetAt(0).getRow(1).getCell(1).getStringCellValue());
            assertEquals("备注", workbook.getSheetAt(0).getRow(0).getCell(2).getStringCellValue());
            assertEquals("Alice", workbook.getSheetAt(0).getRow(2).getCell(0).getStringCellValue());
            assertEquals("A1:B1", workbook.getSheetAt(0).getMergedRegion(0).formatAsString());
        }

        List<ComplexHeaderReadDto> rows = Excel.read(new ByteArrayInputStream(outputStream.toByteArray()))
                .fileName("complex.xlsx")
                .list(ComplexHeaderReadDto.class)
                .parse()
                .getList(ComplexHeaderReadDto.class);

        assertEquals(1, rows.size());
        assertEquals("Alice", rows.get(0).getName());
        assertEquals(Integer.valueOf(18), rows.get(0).getAge());
        assertEquals("OK", rows.get(0).getRemark());
    }

    @Test
    public void readsDynamicMapWithComplexHeaders() {
        ByteArrayOutputStream outputStream = new ByteArrayOutputStream();

        Excel.write(outputStream)
                .list(Arrays.asList(new ComplexHeaderDto("Alice", 18, "OK")), ComplexHeaderDto.class)
                .export();

        List<Map<String, Object>> rows = Excel.read(new ByteArrayInputStream(outputStream.toByteArray()))
                .fileName("complex.xlsx")
                .headRowNumber(0)
                .headRowCount(2)
                .mapList();

        assertEquals(1, rows.size());
        assertEquals("Alice", rows.get(0).get("基础信息.姓名"));
        assertEquals("18", rows.get(0).get("基础信息.年龄"));
        assertEquals("OK", rows.get(0).get("备注"));
    }

    @Test
    public void exportsRichTextValue() throws Exception {
        ByteArrayOutputStream outputStream = new ByteArrayOutputStream();

        Excel.write(outputStream)
                .list(Arrays.asList(new RichTextDto(ExcelRichTextValue.of("hello rich text"))), RichTextDto.class)
                .export();

        try (Workbook workbook = WorkbookFactory.create(new ByteArrayInputStream(outputStream.toByteArray()))) {
            assertEquals("hello rich text", workbook.getSheetAt(0).getRow(1).getCell(0).getStringCellValue());
        }
    }

    @Test
    public void readsHorizontalModelList() throws Exception {
        byte[] workbookBytes;
        try (XSSFWorkbook workbook = new XSSFWorkbook();
             ByteArrayOutputStream outputStream = new ByteArrayOutputStream()) {
            XSSFSheet sheet = workbook.createSheet("horizontal");
            sheet.createRow(0).createCell(0).setCellValue("姓名");
            sheet.getRow(0).createCell(1).setCellValue("Alice");
            sheet.getRow(0).createCell(2).setCellValue("Bob");
            sheet.createRow(1).createCell(0).setCellValue("年龄");
            sheet.getRow(1).createCell(1).setCellValue("18");
            sheet.getRow(1).createCell(2).setCellValue("20");
            workbook.write(outputStream);
            workbookBytes = outputStream.toByteArray();
        }

        List<CsvDto> rows = Excel.read(new ByteArrayInputStream(workbookBytes))
                .fileName("horizontal.xlsx")
                .sheet("horizontal")
                .horizontalList(CsvDto.class);

        assertEquals(2, rows.size());
        assertEquals("Alice", rows.get(0).getName());
        assertEquals(Integer.valueOf(20), rows.get(1).getAge());
    }

    @Test
    public void readsVerticalRepeatedList() throws Exception {
        byte[] workbookBytes;
        try (XSSFWorkbook workbook = new XSSFWorkbook();
             ByteArrayOutputStream outputStream = new ByteArrayOutputStream()) {
            XSSFSheet sheet = workbook.createSheet("vertical");
            sheet.createRow(0).createCell(0).setCellValue("姓名");
            sheet.getRow(0).createCell(1).setCellValue("Alice");
            sheet.createRow(1).createCell(0).setCellValue("年龄");
            sheet.getRow(1).createCell(1).setCellValue("18");
            sheet.createRow(2).createCell(0).setCellValue("姓名");
            sheet.getRow(2).createCell(1).setCellValue("Bob");
            sheet.createRow(3).createCell(0).setCellValue("年龄");
            sheet.getRow(3).createCell(1).setCellValue("20");
            workbook.write(outputStream);
            workbookBytes = outputStream.toByteArray();
        }

        List<CsvDto> rows = Excel.read(new ByteArrayInputStream(workbookBytes))
                .fileName("vertical.xlsx")
                .sheet("vertical")
                .verticalList(CsvDto.class);

        assertEquals(2, rows.size());
        assertEquals("Alice", rows.get(0).getName());
        assertEquals(Integer.valueOf(20), rows.get(1).getAge());
    }

    @Test
    public void readsWorkbookHeadersCommentsHyperlinksAndMergedCells() throws Exception {
        byte[] workbookBytes = createWorkbookWithMetadata();

        List<ExcelHeaderInfo> headers = Excel.read(new ByteArrayInputStream(workbookBytes))
                .fileName("metadata.xlsx")
                .headRowNumber(1)
                .readHeaders();
        assertEquals(2, headers.size());
        assertEquals("姓名", headers.get(0).getTitle());
        assertEquals(Integer.valueOf(1), headers.get(0).getRowIndex());

        List<ExcelCellComment> comments = Excel.read(new ByteArrayInputStream(workbookBytes))
                .fileName("metadata.xlsx")
                .readComments();
        assertEquals(1, comments.size());
        assertEquals("必填", comments.get(0).getText());
        assertEquals("system", comments.get(0).getAuthor());

        List<ExcelCellHyperlink> hyperlinks = Excel.read(new ByteArrayInputStream(workbookBytes))
                .fileName("metadata.xlsx")
                .readHyperlinks();
        assertEquals(1, hyperlinks.size());
        assertEquals("https://github.com/wyatt-wang/flux-excel", hyperlinks.get(0).getAddress());

        List<ExcelMergedCell> mergedCells = Excel.read(new ByteArrayInputStream(workbookBytes))
                .fileName("metadata.xlsx")
                .readMergedCells();
        assertEquals(1, mergedCells.size());
        assertEquals("A3:A4", mergedCells.get(0).getCellRange());
        assertEquals("Alice", mergedCells.get(0).getValue());
    }

    @Test
    public void readsDynamicHeadersAndMergedTableData() throws Exception {
        byte[] workbookBytes = createWorkbookWithMetadata();

        List<String> headers = Excel.read(new ByteArrayInputStream(workbookBytes))
                .fileName("metadata.xlsx")
                .headRowNumber(1)
                .headers();
        assertEquals(Arrays.asList("姓名", "主页"), headers);

        List<Map<String, Object>> rows = Excel.read(new ByteArrayInputStream(workbookBytes))
                .fileName("metadata.xlsx")
                .headRowNumber(1)
                .dataStartRow(2)
                .mapList();
        assertEquals(2, rows.size());
        assertEquals("Alice", rows.get(0).get("姓名"));
        assertEquals("Alice", rows.get(1).get("姓名"));
        assertEquals("20", rows.get(1).get("主页"));
    }

    @Test
    public void readsMergedCellsIntoAnnotatedModel() throws Exception {
        byte[] workbookBytes = createWorkbookWithMetadata();

        List<MergedReadDto> rows = Excel.read(new ByteArrayInputStream(workbookBytes))
                .fileName("metadata.xlsx")
                .list(MergedReadDto.class)
                .parse()
                .getList(MergedReadDto.class);

        assertEquals(2, rows.size());
        assertEquals("Alice", rows.get(0).getName());
        assertEquals("Alice", rows.get(1).getName());
        assertEquals("20", rows.get(1).getHome());
    }

    private byte[] createWorkbookWithMetadata() throws Exception {
        try (XSSFWorkbook workbook = new XSSFWorkbook();
             ByteArrayOutputStream outputStream = new ByteArrayOutputStream()) {
            XSSFSheet sheet = workbook.createSheet("metadata");
            Row header = sheet.createRow(1);
            header.createCell(0).setCellValue("姓名");
            header.createCell(1).setCellValue("主页");

            Row firstDataRow = sheet.createRow(2);
            Cell nameCell = firstDataRow.createCell(0);
            nameCell.setCellValue("Alice");
            Cell linkCell = firstDataRow.createCell(1);
            linkCell.setCellValue("flux-excel");

            Row secondDataRow = sheet.createRow(3);
            secondDataRow.createCell(1).setCellValue("20");
            sheet.addMergedRegion(new CellRangeAddress(2, 3, 0, 0));

            CreationHelper creationHelper = workbook.getCreationHelper();
            Drawing<?> drawing = sheet.createDrawingPatriarch();
            ClientAnchor anchor = creationHelper.createClientAnchor();
            anchor.setCol1(0);
            anchor.setCol2(2);
            anchor.setRow1(1);
            anchor.setRow2(3);
            Comment comment = drawing.createCellComment(anchor);
            comment.setAuthor("system");
            comment.setString(creationHelper.createRichTextString("必填"));
            header.getCell(0).setCellComment(comment);

            Hyperlink hyperlink = creationHelper.createHyperlink(HyperlinkType.URL);
            hyperlink.setAddress("https://github.com/wyatt-wang/flux-excel");
            linkCell.setHyperlink(hyperlink);

            workbook.write(outputStream);
            return outputStream.toByteArray();
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

    @ExcelWrite
    public static class ComplexHeaderDto extends ExcelBaseModel {
        @ExcelWriteProperty(titleName = "姓名", head = {"基础信息", "姓名"})
        private String name;

        @ExcelWriteProperty(titleName = "年龄", head = {"基础信息", "年龄"})
        private Integer age;

        @ExcelWriteProperty(titleName = "备注")
        private String remark;

        public ComplexHeaderDto() {
        }

        public ComplexHeaderDto(String name, Integer age, String remark) {
            this.name = name;
            this.age = age;
            this.remark = remark;
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

        public String getRemark() {
            return remark;
        }

        public void setRemark(String remark) {
            this.remark = remark;
        }
    }

    @ExcelRead
    public static class ComplexHeaderReadDto extends ExcelBaseModel {
        @ExcelReadProperty(titleName = "姓名", head = {"基础信息", "姓名"})
        private String name;

        @ExcelReadProperty(titleName = "年龄", head = {"基础信息", "年龄"})
        private Integer age;

        @ExcelReadProperty(titleName = "备注")
        private String remark;

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

        public String getRemark() {
            return remark;
        }

        public void setRemark(String remark) {
            this.remark = remark;
        }
    }

    @ExcelWrite
    public static class RichTextDto extends ExcelBaseModel {
        @ExcelWriteProperty(titleName = "富文本")
        private ExcelRichTextValue text;

        public RichTextDto() {
        }

        public RichTextDto(ExcelRichTextValue text) {
            this.text = text;
        }

        public ExcelRichTextValue getText() {
            return text;
        }

        public void setText(ExcelRichTextValue text) {
            this.text = text;
        }
    }

    @ExcelRead
    public static class MergedReadDto extends ExcelBaseModel {
        @ExcelReadProperty(titleName = "姓名")
        private String name;

        @ExcelReadProperty(titleName = "主页")
        private String home;

        public String getName() {
            return name;
        }

        public void setName(String name) {
            this.name = name;
        }

        public String getHome() {
            return home;
        }

        public void setHome(String home) {
            this.home = home;
        }
    }
}
