# flux-excel

[中文文档](README.zh-CN.md)

flux-excel is a Java Excel import and export toolkit built on Apache POI. It is designed to cover common Excel workflows with a small set of annotations and fluent APIs, including standard import/export, template filling, large data export, batch zip export, event-based reading, field formatting, comments, dropdown validation, image reading/writing, and related scenarios.

The project contains two modules:

- `flux-excel-core`: core import and export features.
- `flux-excel-spring-boot-starter`: Spring Boot auto-configuration, including template directory resource resolution. Model and template metadata are loaded lazily by default from the entry point that uses them.

## Features

- Fluent entry points: `Excel.write(...)`, `Excel.read(...)`, `Excel.largeWrite(...)`, `Excel.batchWrite(...)`, and `Excel.eventRead(...)`.
- Annotation-driven models: use `@ExcelWrite`, `@ExcelWriteProperty`, `@ExcelRead`, and `@ExcelReadProperty` to define headers, column width, formatting, styles, validation, and more.
- Import/export formatting for `java.util.Date`, `Calendar`, and common `java.time` types.
- Supports regular objects, lists, and dynamic `List<Map<String, Object>>` tables.
- Supports regular Workbook reading and event-based reading.
- Supports reading headers.
- Supports reading comments.
- Supports reading hyperlinks.
- Supports reading merged cells.
- Supports dynamic headers and table data.
- Supports `.xls`, `.xlsx`, and `.csv`.
- Supports template export and template validation.
- Supports text watermarks and custom headers for empty exports.
- Supports large list export, sheet splitting, multi-threaded export, and batch zip export.
- Supports Spring Boot auto-configuration.

## Requirements

- JDK 21+
- Maven 3.8+
- Spring Boot 3.x, only when using `flux-excel-spring-boot-starter`

## Installation

If the dependency has been published to a Maven repository, add it directly:

```xml
<dependency>
    <groupId>com.github</groupId>
    <artifactId>flux-excel-core</artifactId>
    <version>1.0.0.RELEASE</version>
</dependency>
```

For Spring Boot projects, the starter is recommended:

```xml
<dependency>
    <groupId>com.github</groupId>
    <artifactId>flux-excel-spring-boot-starter</artifactId>
    <version>1.0.0.RELEASE</version>
</dependency>
```

If the artifacts have not been published yet, install them locally from source first:

```bash
mvn clean install
```

Then reference the coordinates above from your application.

## Quick Start

### 1. Define an export/import model

Models should extend `ExcelBaseModel`.

```java
import com.github.excel.annotation.ExcelRead;
import com.github.excel.annotation.ExcelReadProperty;
import com.github.excel.annotation.ExcelWrite;
import com.github.excel.annotation.ExcelWriteProperty;
import com.github.excel.model.ExcelBaseModel;
import lombok.Data;

import java.math.BigDecimal;
import java.time.LocalDateTime;

@Data
@ExcelWrite
@ExcelRead(checkTitle = true)
public class UserExcelDTO extends ExcelBaseModel {

    @ExcelWriteProperty(titleName = "姓名", colWidth = 300)
    @ExcelReadProperty(titleName = "姓名", checkNull = true)
    private String name;

    @ExcelWriteProperty(titleName = "年龄")
    @ExcelReadProperty(titleName = "年龄")
    private Integer age;

    @ExcelWriteProperty(titleName = "余额", formatPattern = "#,##0.00")
    @ExcelReadProperty(titleName = "余额")
    private BigDecimal balance;

    @ExcelWriteProperty(titleName = "创建时间", formatPattern = "yyyy-MM-dd HH:mm:ss")
    @ExcelReadProperty(titleName = "创建时间", formatPattern = "yyyy-MM-dd HH:mm:ss")
    private LocalDateTime createTime;
}
```

## Fluent API Overview

Common entry points are available from `com.github.excel.Excel`:

| Entry | Description |
| --- | --- |
| `Excel.write(OutputStream)` | Export to an output stream |
| `Excel.write(File)` | Export to a file |
| `Excel.write(HttpServletRequest, HttpServletResponse)` | Export to an HTTP response |
| `Excel.read(InputStream)` | Read from an input stream |
| `Excel.read(File)` | Read from a file |
| `Excel.largeWrite(OutputStream)` | SXSSF large data export to an output stream |
| `Excel.largeWrite(File)` | SXSSF large data export to a file |
| `Excel.batchWrite(String outputDirPath)` | Batch export multiple files and compress them |
| `Excel.eventRead(InputStream)` | Event-based reading from an input stream |
| `Excel.eventRead(File)` | Event-based reading from a file |

Standard export supports `.xlsx`, `.xls`, and `.csv`. Large data export is based on `SXSSFWorkbook` and only generates `.xlsx`. Event-based reading is suitable for large imports because it does not create a full `Workbook`.

### 2. Export Excel

```java
import com.github.excel.Excel;
import com.github.excel.enums.ExcelSuffixEnum;

import java.io.FileOutputStream;
import java.io.OutputStream;
import java.util.List;

List<UserExcelDTO> users = queryUsers();

try (OutputStream outputStream = new FileOutputStream("users.xlsx")) {
    Excel.write(outputStream)
            .fileName("users.xlsx")
            .suffix(ExcelSuffixEnum.XLSX)
            .sheet("用户")
            .list(users, UserExcelDTO.class)
            .export();
}
```

### 3. Read Excel

```java
import com.github.excel.Excel;

import java.io.FileInputStream;
import java.io.InputStream;
import java.util.List;

try (InputStream inputStream = new FileInputStream("users.xlsx")) {
    List<UserExcelDTO> users = Excel.read(inputStream)
            .fileName("users.xlsx")
            .sheet(0)
            .list(UserExcelDTO.class)
            .parse()
            .getList(UserExcelDTO.class);
}
```

## Export Usage

### Export a single object

```java
UserExcelDTO user = queryUser();

Excel.write(new File("user.xlsx"))
        .fileName("user.xlsx")
        .suffix(ExcelSuffixEnum.XLSX)
        .sheet("用户")
        .model(user, UserExcelDTO.class)
        .export();
```

### Specify the start position

`at(rowIndex, colIndex)` uses zero-based row and column indexes.

```java
Excel.write(outputStream)
        .fileName("users.xlsx")
        .suffix(ExcelSuffixEnum.XLSX)
        .sheet("用户")
        .at(2, 1)
        .list(users, UserExcelDTO.class)
        .export();
```

### Exclude fields

`exclude(...)` applies to the most recent data model added by `model(...)` or `list(...)`.

```java
Excel.write(outputStream)
        .fileName("users.xlsx")
        .suffix(ExcelSuffixEnum.XLSX)
        .sheet("用户")
        .list(users, UserExcelDTO.class)
        .exclude("balance")
        .export();
```

You can also pass excluded fields directly when adding a list. This is supported by large data export as well:

```java
Excel.largeWrite(outputStream)
        .sheet("用户")
        .list(users, UserExcelDTO.class, "balance")
        .export();
```

### Export with a template

Template files should be placed in the configured template directory. Plain Java projects can load templates manually. Spring Boot projects can use auto-configuration, as shown later.

```java
Excel.write(outputStream)
        .fileName("report.xlsx")
        .suffix(ExcelSuffixEnum.XLSX)
        .template("report-template.xlsx")
        .sheet("报告")
        .fillTemplate(true)
        .list(rows, ReportRowDTO.class)
        .export();
```

You can also use a template from disk:

```java
Excel.write(outputStream)
        .fileName("report.xlsx")
        .suffix(ExcelSuffixEnum.XLSX)
        .templateFilePath("/data/templates/report-template.xlsx")
        .templatePassword("optional-password")
        .sheet("报告")
        .list(rows, ReportRowDTO.class)
        .export();
```

### Custom write processing

Use `custom(...)` to access the underlying POI `Workbook` for additional processing.

```java
Excel.write(outputStream)
        .fileName("users.xlsx")
        .suffix(ExcelSuffixEnum.XLSX)
        .sheet("用户")
        .list(users, UserExcelDTO.class)
        .custom(workbook -> {
            workbook.createSheet("说明")
                    .createRow(0)
                    .createCell(0)
                    .setCellValue("Generated by flux-excel");
        })
        .export();
```

### Write cells and merged cells

`column(...)` writes a single cell. `mergeColumn(...)` merges a region and writes a value. Row and column indexes are zero-based.

```java
ExcelWriterCellParam title = new ExcelWriterCellParam()
        .setSheetName("汇总")
        .setRowIndex(0)
        .setColIndex(0)
        .setValue("用户汇总");

ExcelWriterMergeParam mergedTitle = new ExcelWriterMergeParam()
        .setSheetName("汇总")
        .setRowIndex(1)
        .setColIndex(0)
        .setEndRowIndex(1)
        .setEndColIndex(3)
        .setValue("基础信息");

Excel.write(outputStream)
        .fileName("summary.xlsx")
        .suffix(ExcelSuffixEnum.XLSX)
        .column(title)
        .mergeColumn(mergedTitle)
        .export();
```

### Comments and data validation

`comment(...)` and `validation(...)` apply to fields in the current export model by field name.

```java
ExcelWriterCommentParam comment = new ExcelWriterCommentParam()
        .setCommentText("必填")
        .setAuthor("system");

ExcelWriterComboParam combo = new ExcelWriterComboParam()
        .setOptions(new String[]{"男", "女"})
        .setCommentText("请选择下拉选项");

ExcelWriterNumberScopeParam ageScope = new ExcelWriterNumberScopeParam()
        .setStart("0")
        .setEnd("120")
        .setCommentText("年龄范围为 0-120");

Excel.write(outputStream)
        .fileName("users.xlsx")
        .suffix(ExcelSuffixEnum.XLSX)
        .sheet("用户")
        .list(users, UserExcelDTO.class)
        .comment("name", comment)
        .validation("sex", combo)
        .validation("age", ageScope)
        .export();
```

### Styles

Register a custom style class with `style(...)`, then reference the style name from annotations.

```java
Excel.write(outputStream)
        .fileName("users.xlsx")
        .suffix(ExcelSuffixEnum.XLSX)
        .style(UserExcelStyle.class)
        .sheet("用户")
        .list(users, UserExcelDTO.class)
        .export();
```

### Dynamic Map/CSV and empty headers

`listMap(...)` is suitable for dynamic tables whose fields are not fixed until runtime. Enable CSV output with `suffix(ExcelSuffixEnum.CSV)`.

```java
List<Map<String, Object>> rows = queryDynamicRows();

Excel.write(outputStream)
        .fileName("users.csv")
        .suffix(ExcelSuffixEnum.CSV)
        .listMap(rows)
        .export();
```

When exporting empty data, you can pass Map headers directly or pass a model class annotated with `@ExcelWriteProperty`.

```java
Map<String, String> headers = new LinkedHashMap<>();
headers.put("name", "姓名");
headers.put("age", "年龄");

Excel.write(outputStream)
        .emptyHeader(headers)
        .export();

Excel.write(outputStream)
        .emptyHeader(UserExcelDTO.class)
        .export();
```

### Text watermark

`watermark(...)` adds a text background watermark to `.xlsx` export files.

```java
Excel.write(outputStream)
        .fileName("users.xlsx")
        .suffix(ExcelSuffixEnum.XLSX)
        .watermark("DRAFT")
        .list(users, UserExcelDTO.class)
        .export();
```

### SXSSF streaming export configuration

For regular `.xlsx` export, call `streaming(true)` to enable `SXSSFWorkbook`. Large data export uses SXSSF by default and supports the same row-window and temporary-file options.

```java
Excel.write(outputStream)
        .fileName("users.xlsx")
        .suffix(ExcelSuffixEnum.XLSX)
        .streaming(true)
        .rowAccessWindowSize(5000)
        .compressTempFiles(true)
        .useSharedStringsTable(false)
        .sheet("用户")
        .list(users, UserExcelDTO.class)
        .export();
```

Configuration:

| Method | Description |
| --- | --- |
| `streaming(true)` | Enable SXSSF for regular `.xlsx` export. `.xls` does not support SXSSF |
| `rowAccessWindowSize(int)` | Row window retained in memory by SXSSF. Default is `10000`; `-1` means rows are not flushed automatically |
| `compressTempFiles(boolean)` | Whether to compress SXSSF temporary files. This reduces disk usage but increases CPU cost |
| `useSharedStringsTable(boolean)` | Whether to use a shared strings table. It improves compatibility but may increase memory usage |

### Export to an HTTP response

```java
@GetMapping("/users/export")
public void exportUsers(HttpServletRequest request, HttpServletResponse response) {
    List<UserExcelDTO> users = userService.queryUsers();

    Excel.write(request, response)
            .fileName("users")
            .suffix(ExcelSuffixEnum.XLSX)
            .sheet("用户")
            .list(users, UserExcelDTO.class)
            .export();
}
```

## Import Usage

### Import entry point and parameters

```java
Excel.read(new File("users.xlsx"))
        .password("optional-password")
        .readOnly(true)
        .readPicture(false)
        .failFast(true)
        .sheet(0)
        .list(UserExcelDTO.class)
        .parse();
```

Common methods:

| Method | Description |
| --- | --- |
| `fileName(String)` | Specify the file name, mainly for detecting suffixes such as `.csv` |
| `template(String)` | Specify an import template name, used with `@ExcelRead(checkTitle = true)` for title validation |
| `password(String)` | Read an encrypted Excel file |
| `closeInputStream(boolean)` | Whether the framework should close the `InputStream` after reading |
| `readPicture(boolean)` | Whether to parse images |
| `readOnly(boolean)` | File read mode |
| `failFast(boolean)` | Whether validation should fail fast |
| `sheet(int)` | Select a sheet by zero-based index |
| `sheet(String)` | Select a sheet by name |
| `headRowNumber(int)` | Header start row or title column, zero-based |
| `headRowCount(int)` | Number of header rows for dynamic multi-header map reading |
| `dataStartRow(int)` | Data start row, or value start column for horizontal reading |
| `dataEndRow(int)` | Data end row, or value end column for horizontal reading |
| `maxRows(int)` | Maximum rows/models to read |

### Read a single object

```java
UserExcelDTO user = Excel.read(new File("user.xlsx"))
        .sheet(0)
        .model(UserExcelDTO.class)
        .parse()
        .getModel(UserExcelDTO.class);
```

### Read a list and close the input stream

```java
List<UserExcelDTO> users = Excel.read(inputStream)
        .fileName("users.xlsx")
        .closeInputStream(true)
        .sheet(0)
        .list(UserExcelDTO.class)
        .parse()
        .getList(UserExcelDTO.class);
```

### Read multiple models or multiple sheets

Call `sheet(...)`, `model(...)`, and `list(...)` repeatedly to register multiple read tasks.

```java
ExcelReadBuilder builder = Excel.read(inputStream)
        .fileName("report.xlsx")
        .sheet(0)
        .model(ReportTitleDTO.class)
        .sheet(1)
        .list(UserExcelDTO.class)
        .parse();

ReportTitleDTO title = builder.getModel(ReportTitleDTO.class);
List<UserExcelDTO> users = builder.getList(UserExcelDTO.class);
```

### Row callback

```java
Excel.read(inputStream)
        .fileName("users.xlsx")
        .sheet(0)
        .list(UserExcelDTO.class, user -> {
            // Called after each row is parsed.
            System.out.println(user.getName());
        })
        .parse();
```

### Batch callback

This is useful when importing a large amount of data and persisting it in batches.

```java
Excel.read(inputStream)
        .fileName("users.xlsx")
        .sheet(0)
        .list(UserExcelDTO.class, new ExcelReaderBatchProcess<UserExcelDTO>() {
            @Override
            public int getBatchSize() {
                return 1000;
            }

            @Override
            public void process(List<UserExcelDTO> dataList) {
                userService.saveBatch(dataList);
            }
        })
        .parse();
```

### Read images

After enabling `readPicture(true)`, model fields can receive cell images as `ExcelReaderPictureModel`, `List<ExcelReaderPictureModel>`, or `byte[]`.

```java
@ExcelRead
public class UserImportDTO extends ExcelBaseModel {
    @ExcelReadProperty(titleName = "头像")
    private List<ExcelReaderPictureModel> avatar;
}

List<UserImportDTO> users = Excel.read(inputStream)
        .fileName("users.xlsx")
        .readPicture(true)
        .sheet(0)
        .list(UserImportDTO.class)
        .parse()
        .getList(UserImportDTO.class);
```

### Custom read processing

```java
Excel.read(inputStream)
        .fileName("users.xlsx")
        .custom(workbook -> {
            String firstSheetName = workbook.getSheetAt(0).getSheetName();
            System.out.println(firstSheetName);
        })
        .sheet(0)
        .list(UserExcelDTO.class)
        .parse();
```

### Read Map/CSV

`mapList()` treats the first row as the header and converts the remaining rows to `List<Map<String, Object>>`.

```java
List<Map<String, Object>> rows = Excel.read(inputStream)
        .fileName("users.csv")
        .mapList();
```

CSV can also be read into a model:

```java
List<UserExcelDTO> users = Excel.read(inputStream)
        .fileName("users.csv")
        .list(UserExcelDTO.class)
        .parse()
        .getList(UserExcelDTO.class);
```

### Complex Headers

Multi-level headers can use the fast-excel/EasyExcel-style `head` or `value` annotation attribute. `titleName` remains supported and is used as the leaf title when `head` is not set.

```java
@ExcelWrite
@ExcelRead
public class UserExcelDTO extends ExcelBaseModel {
    @ExcelWriteProperty(titleName = "姓名", head = {"基础信息", "姓名"})
    @ExcelReadProperty(titleName = "姓名", head = {"基础信息", "姓名"})
    private String name;

    @ExcelWriteProperty(titleName = "年龄", head = {"基础信息", "年龄"})
    @ExcelReadProperty(titleName = "年龄", head = {"基础信息", "年龄"})
    private Integer age;
}
```

For dynamic data, combine multiple header rows with `headRowCount(...)`:

```java
List<Map<String, Object>> rows = Excel.read(inputStream)
        .fileName("users.xlsx")
        .headRowNumber(0)
        .headRowCount(2)
        .mapList();
```

### Horizontal And Vertical Reading

Horizontal reading treats the first column as field titles and each following column as one model. Vertical repeated-list reading treats a title/value block as one model and starts a new model when a field title repeats.

```java
List<UserExcelDTO> horizontalUsers = Excel.read(inputStream)
        .fileName("users.xlsx")
        .horizontalList(UserExcelDTO.class);

List<UserExcelDTO> verticalUsers = Excel.read(inputStream)
        .fileName("users.xlsx")
        .verticalList(UserExcelDTO.class);
```

For a form plus list workbook, register both tasks on the same builder:

```java
ExcelReadBuilder builder = Excel.read(inputStream)
        .fileName("report.xlsx")
        .model(ReportTitleDTO.class)
        .list(UserExcelDTO.class)
        .parse();
```

### Rich Text Export

Cells can receive Apache POI `RichTextString` directly, or the convenience wrapper `ExcelRichTextValue`.

```java
@ExcelWrite
public class RichTextDTO extends ExcelBaseModel {
    @ExcelWriteProperty(titleName = "说明")
    private ExcelRichTextValue description;
}
```

## Large Data Export

Large data export uses `Excel.largeWrite(...)`. This mode reuses the internal large data writer and is suitable for million-row lists or scenarios that need to split sheets by a maximum row count.

```java
try (OutputStream outputStream = new FileOutputStream("large-users.xlsx")) {
    Excel.largeWrite(outputStream)
            .sheet("用户")
            .rowAccessWindowSize(5000)
            .compressTempFiles(true)
            .sheetRowMaxCount(100000)
            .modelClass(UserExcelDTO.class)
            .list(users)
            .export()
            .close();
}
```

The same large data export file can append multiple lists. Each `list(...)` call can specify a model and excluded fields.

```java
Excel.largeWrite(outputStream)
        .sheet("用户")
        .sheetRowMaxCount(100000)
        .list(users, UserExcelDTO.class, "balance")
        .list(orders, OrderExcelDTO.class)
        .export()
        .close();
```

## Batch Export Zip

Batch export concurrently generates multiple Excel files and finally compresses them into a zip. `ExcelBatchWriteBuilder(String outputDirPath, int maxPoolSize)` can be used to specify the maximum thread count.

```java
Excel.batchWrite("/tmp/flux-excel")
        .rowAccessWindowSize(5000)
        .compressTempFiles(true)
        .sheet("用户")
        .file("users-1")
        .list(users1, UserExcelDTO.class)
        .file("users-2")
        .list(users2, UserExcelDTO.class)
        .exportZip("users");
```

Export a zip to an HTTP response:

```java
new ExcelBatchWriteBuilder("/tmp/flux-excel", 4)
        .file("users")
        .sheet("用户")
        .list(users, UserExcelDTO.class)
        .export(request, response, "users");
```

## Event Read

Event-based reading is suitable for parsing very large files. Provide a row reader and a batch handler when using it.

```java
Excel.eventRead(inputStream)
        .fileName("large-users.xlsx")
        .rowReader((sheetIndex, curRow, rowValue) -> {
            // Convert the raw row to a business object.
            return convertRow(sheetIndex, curRow, rowValue);
        })
        .batchHandler(new AbstractEventBatchHandler<UserExcelDTO>(1000) {
            @Override
            public void batchExecute(List<UserExcelDTO> dataList) {
                userService.saveBatch(dataList);
            }
        })
        .parse();
```

`rowReader` receives a zero-based `sheetIndex`, the original Excel row number as `curRow`, and the current row as a list of strings in `rowValue`. Event-based reading is suitable for large imports, but it does not execute the model annotation mapping, image reading, or template validation used by regular `Excel.read(...)`.

## Compatible Writer API

Prefer the fluent API for new code. When compatibility with legacy calls is required, use `Excel.write(...).writer()` to access the underlying `ExcelWriter`.

```java
ExcelWriter writer = Excel.write(outputStream)
        .suffix(ExcelSuffixEnum.XLSX)
        .writer();

writer.setStreaming(true)
        .addModelList(users, "用户", true, UserExcelDTO.class)
        .addValidationOrComment("age", ageScope)
        .process(outputStream, "users.xlsx", ExcelSuffixEnum.XLSX);
```

Common compatible methods:

| Method | Description |
| --- | --- |
| `addModel(...)` | Add a single model |
| `addModelList(...)` | Add a model list |
| `addCustomColumn(...)` | Add a custom cell |
| `addMergeCustomColumn(...)` | Add a merged cell |
| `addValidationOrComment(...)` | Add a comment, numeric range validation, or dropdown validation |
| `setStreaming(boolean)` | Set whether regular `.xlsx` export uses SXSSF |
| `selectSheet(String)` | Set the sheet selected when the file is opened |
| `process(...)` / `export(...)` | Write to a stream, file, or response |

## Spring Boot Starter

After adding the starter, application startup automatically:

- Registers resource resolution for export and import template directories.
- Registers the optional import template exclusion hook.
- Registers the default `ExcelLargeListWriter` bean.

DTO metadata and template contents are not scanned or parsed at startup. They are loaded lazily and cached when a corresponding `Class` or template name is first used by import/export.

Configuration example:

```yaml
excel:
  export-template-dir: excel-template
  import-template-dir: import-excel-template
```

Resource directory example:

```text
src/main/resources/
  excel-template/
    report-template.xlsx
  import-excel-template/
    users-import-template.xlsx
```

## Annotations

### `@ExcelWrite`

Marks a model class as exportable. Common attributes:

| Attribute | Description |
| --- | --- |
| `fillTitle` | Whether to fill headers automatically |
| `freezeTitle` | Whether to freeze the header row |
| `rowIndex` / `colIndex` | Default start row and column |
| `fillStyle` | Fill direction |
| `titleStyleName` / `contentStyleName` | Header and content style names |
| `fillType` | List fill mode |
| `themeName` / `customThemeName` | Theme configuration |
| `incrementSequenceNo` | Whether to fill incremental sequence numbers |
| `filterTitle` | Whether to enable header filtering |

### `@ExcelWriteProperty`

Marks an export field. Common attributes:

| Attribute | Description |
| --- | --- |
| `titleName` | Header name |
| `head` / `value` | Multi-level header path, compatible with fast-excel/EasyExcel style |
| `rowHeight` / `colWidth` | Row height and column width |
| `rowIndex` / `colIndex` | Fixed row and column for the field |
| `formatPattern` | Output format |
| `formatter` | Custom export formatter |
| `commentText` | Comment content |
| `dropDownOptions` | Dropdown options |
| `disable` | Whether to disable this field |
| `index` | Sort order |

### `@ExcelRead`

Marks a model class as importable. Common attributes:

| Attribute | Description |
| --- | --- |
| `checkTitle` | Whether to validate headers |
| `enableSeparator` | Whether to enable separator parsing |

### `@ExcelReadProperty`

Marks an import field. Common attributes:

| Attribute | Description |
| --- | --- |
| `titleName` | Matching header name |
| `head` / `value` | Multi-level header path, compatible with fast-excel/EasyExcel style |
| `separator` | Separator |
| `formatPattern` | Read format |
| `formatter` | Custom read formatter |
| `checkNull` | Whether to validate non-null values |
| `disable` | Whether to disable this field |

## Custom Formatters

### Export formatter

```java
public class MoneyWriterFormat implements ExcelWriterDataFormat {
    @Override
    public Object format(Object value, String pattern) {
        if (value == null) {
            return null;
        }
        return "￥" + value;
    }
}
```

```java
@ExcelWriteProperty(titleName = "金额", formatter = MoneyWriterFormat.class)
private BigDecimal amount;
```

### Import formatter

```java
public class MoneyReaderFormat implements ExcelReaderDataFormat {
    @Override
    public Object format(Object value, String pattern, Class<?> fieldType) {
        if (value == null) {
            return null;
        }
        return new BigDecimal(value.toString().replace("￥", ""));
    }
}
```

```java
@ExcelReadProperty(titleName = "金额", formatter = MoneyReaderFormat.class)
private BigDecimal amount;
```

## Template Loading Without Spring Boot

Non-Spring Boot projects also lazily load model metadata by the provided model `Class`. Templates can be loaded lazily by template name after configuring directories with `ExcelMetadataRegistry`, or you can continue to proactively warm them through the legacy API:

```java
ExcelMetadataRegistry.configureTemplateDirectory("/path/to/export/templates", "/path/to/import/templates", null);

// Optional warm-up, not required.
ExcelMetadataRegistry.loadExcelTemplate("/path/to/export/templates");
ExcelMetadataRegistry.loadImportExcelTemplate("/path/to/import/templates", null);
```

To explicitly reload metadata, clear the cache. The next import/export will parse it again:

```java
ExcelMetadataRegistry.clearModel(UserExcelDTO.class);
ExcelMetadataRegistry.clearTemplate("user-template.xlsx");
```

`ExcelBootLoader` is still retained as a compatibility facade. Legacy calls are delegated to `ExcelMetadataRegistry`.

## Build And Test

Compile:

```bash
mvn -pl flux-excel-core -am compile
```

Run tests:

```bash
mvn -pl flux-excel-core test
```

Full build:

```bash
mvn clean install
```

## Project Structure

```text
flux-excel/
  flux-excel-core/                 # Core import/export features
  flux-excel-spring-boot-starter/  # Spring Boot auto-configuration
  pom.xml
  README.md
  README.zh-CN.md
  LICENSE
```

## Roadmap

- Continue splitting `BaseExcelWriter` to reduce the complexity of the core write logic.
- Improve unit tests and large data export performance tests.
- Add more Spring Boot sample projects.

## Contributing

Issues and pull requests are welcome. Before submitting changes, run:

```bash
mvn clean test
```

When submitting a PR, please describe:

- The purpose and usage scenario of the change.
- Whether it has compatibility impact.
- Whether tests were added or updated.

## License

flux-excel is open-sourced under the [Apache License 2.0](LICENSE).
