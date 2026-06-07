# flux-excel

flux-excel 是一个基于 Apache POI 的 Java Excel 导入导出工具库，目标是用少量注解和链式 API 完成常见 Excel 场景：普通导入导出、模板填充、大数据导出、批量导出压缩、事件模式读取、字段格式化、批注、下拉校验、图片读写等。

项目包含两个模块：

- `flux-excel-core`：核心导入导出能力。
- `flux-excel-spring-boot-starter`：Spring Boot 自动配置，支持模板目录资源定位，模型和模板默认按使用入口懒加载。

## Features

- 链式调用入口：`Excel.write(...)`、`Excel.read(...)`、`Excel.largeWrite(...)`、`Excel.batchWrite(...)`、`Excel.eventRead(...)`。
- 注解驱动模型：通过 `@ExcelWrite`、`@ExcelWriteProperty`、`@ExcelRead`、`@ExcelReadProperty` 定义表头、列宽、格式化、样式、校验等。
- 支持 `java.util.Date`、`Calendar` 和常用 `java.time` 类型的导入导出格式化。
- 支持普通对象、列表和 `List<Map<String, Object>>` 动态表导入导出。
- 支持普通 Workbook 读取和事件模式读取。
- 支持 `.xls`、`.xlsx` 和 `.csv`。
- 支持模板导出和模板校验。
- 支持导出文字水印、空数据自定义表头。
- 支持大数据列表导出、分 sheet 导出、多线程导出和批量导出压缩。
- 支持 Spring Boot 自动装配。

## Requirements

- JDK 21+
- Maven 3.8+
- Spring Boot 3.x，只有使用 `flux-excel-spring-boot-starter` 时需要

## Installation

如果依赖已发布到 Maven 仓库，可直接引入：

```xml
<dependency>
    <groupId>com.github</groupId>
    <artifactId>flux-excel-core</artifactId>
    <version>1.0.0.RELEASE</version>
</dependency>
```

Spring Boot 项目推荐使用 starter：

```xml
<dependency>
    <groupId>com.github</groupId>
    <artifactId>flux-excel-spring-boot-starter</artifactId>
    <version>1.0.0.RELEASE</version>
</dependency>
```

如果还没有发布到 Maven 仓库，可以先从源码安装到本地：

```bash
mvn clean install
```

然后在业务项目中按上面的坐标引用。

## Quick Start

### 1. 定义导出和导入模型

模型需要继承 `ExcelBaseModel`。

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

常用入口都在 `com.github.excel.Excel`：

| 入口 | 说明 |
| --- | --- |
| `Excel.write(OutputStream)` | 导出到输出流 |
| `Excel.write(File)` | 导出到文件 |
| `Excel.write(HttpServletRequest, HttpServletResponse)` | 导出到 HTTP 响应 |
| `Excel.read(InputStream)` | 从输入流读取 |
| `Excel.read(File)` | 从文件读取 |
| `Excel.largeWrite(OutputStream)` | SXSSF 大数据导出到输出流 |
| `Excel.largeWrite(File)` | SXSSF 大数据导出到文件 |
| `Excel.batchWrite(String outputDirPath)` | 多文件批量导出并压缩 |
| `Excel.eventRead(InputStream)` | 事件模式读取输入流 |
| `Excel.eventRead(File)` | 事件模式读取文件 |

普通导出支持 `.xlsx`、`.xls`、`.csv`。大数据导出基于 `SXSSFWorkbook`，只生成 `.xlsx`。事件模式读取适合大文件导入，不会创建完整 `Workbook`。

### 2. 导出 Excel

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

### 3. 读取 Excel

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

### 导出单个对象

```java
UserExcelDTO user = queryUser();

Excel.write(new File("user.xlsx"))
        .fileName("user.xlsx")
        .suffix(ExcelSuffixEnum.XLSX)
        .sheet("用户")
        .model(user, UserExcelDTO.class)
        .export();
```

### 指定起始位置

`at(rowIndex, colIndex)` 使用从 `0` 开始的行列下标。

```java
Excel.write(outputStream)
        .fileName("users.xlsx")
        .suffix(ExcelSuffixEnum.XLSX)
        .sheet("用户")
        .at(2, 1)
        .list(users, UserExcelDTO.class)
        .export();
```

### 排除字段

`exclude(...)` 作用于最近一次 `model(...)` 或 `list(...)` 添加的数据模型。

```java
Excel.write(outputStream)
        .fileName("users.xlsx")
        .suffix(ExcelSuffixEnum.XLSX)
        .sheet("用户")
        .list(users, UserExcelDTO.class)
        .exclude("balance")
        .export();
```

也可以在添加列表时直接传入排除字段，适用于大数据导出：

```java
Excel.largeWrite(outputStream)
        .sheet("用户")
        .list(users, UserExcelDTO.class, "balance")
        .export();
```

### 使用模板导出

模板文件需要提前放到模板目录中。普通 Java 项目可以手动加载模板；Spring Boot 项目可通过配置自动加载，见下文。

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

也可以直接使用磁盘模板：

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

### 自定义写入

可以通过 `custom(...)` 获取底层 POI `Workbook` 做补充处理。

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

### 写入单元格和合并单元格

`column(...)` 用于写入单个单元格，`mergeColumn(...)` 用于合并区域并写入值。行列下标均从 `0` 开始。

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

### 批注和数据校验

`comment(...)`、`validation(...)` 按字段名作用于当前导出的模型字段。

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

### 样式

通过 `style(...)` 注册自定义样式类，然后在注解中使用样式名。

```java
Excel.write(outputStream)
        .fileName("users.xlsx")
        .suffix(ExcelSuffixEnum.XLSX)
        .style(UserExcelStyle.class)
        .sheet("用户")
        .list(users, UserExcelDTO.class)
        .export();
```

### 动态 Map/CSV 和空表头

`listMap(...)` 适合运行时字段不固定的动态表。CSV 通过 `suffix(ExcelSuffixEnum.CSV)` 开启。

```java
List<Map<String, Object>> rows = queryDynamicRows();

Excel.write(outputStream)
        .fileName("users.csv")
        .suffix(ExcelSuffixEnum.CSV)
        .listMap(rows)
        .export();
```

空数据导出时可直接传 Map 表头，也可传已标注 `@ExcelWriteProperty` 的模型类。

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

### 文字水印

`watermark(...)` 会在 `.xlsx` 导出文件中添加文字背景水印。

```java
Excel.write(outputStream)
        .fileName("users.xlsx")
        .suffix(ExcelSuffixEnum.XLSX)
        .watermark("DRAFT")
        .list(users, UserExcelDTO.class)
        .export();
```

### SXSSF 流式导出配置

普通 `.xlsx` 导出可通过 `streaming(true)` 启用 `SXSSFWorkbook`。大数据导出默认就是 SXSSF，也支持相同的窗口和临时文件配置。

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

配置说明：

| 方法 | 说明 |
| --- | --- |
| `streaming(true)` | 普通 `.xlsx` 导出启用 SXSSF；`.xls` 不支持 SXSSF |
| `rowAccessWindowSize(int)` | SXSSF 内存中保留的行窗口，默认 `10000`，`-1` 表示不自动刷出 |
| `compressTempFiles(boolean)` | 是否压缩 SXSSF 临时文件，降低磁盘占用但会增加 CPU 开销 |
| `useSharedStringsTable(boolean)` | 是否使用共享字符串表，兼容性更好但可能增加内存占用 |

### 导出到 HTTP 响应

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

### 导入入口和参数

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

常用方法：

| 方法 | 说明 |
| --- | --- |
| `fileName(String)` | 指定文件名，主要用于判断 `.csv` 等后缀 |
| `template(String)` | 指定导入模板名，配合 `@ExcelRead(checkTitle = true)` 校验 |
| `password(String)` | 读取加密 Excel |
| `closeInputStream(boolean)` | `InputStream` 读取完成后是否由框架关闭 |
| `readPicture(boolean)` | 是否解析图片 |
| `readOnly(boolean)` | 文件读取模式 |
| `failFast(boolean)` | 校验失败时是否快速失败 |
| `sheet(int)` | 选择 sheet，下标从 `0` 开始 |

### 读取单个对象

```java
UserExcelDTO user = Excel.read(new File("user.xlsx"))
        .sheet(0)
        .model(UserExcelDTO.class)
        .parse()
        .getModel(UserExcelDTO.class);
```

### 读取列表并关闭输入流

```java
List<UserExcelDTO> users = Excel.read(inputStream)
        .fileName("users.xlsx")
        .closeInputStream(true)
        .sheet(0)
        .list(UserExcelDTO.class)
        .parse()
        .getList(UserExcelDTO.class);
```

### 同时读取多个模型或多个 Sheet

可以连续调用 `sheet(...)`、`model(...)`、`list(...)` 注册多个读取任务。

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

### 行级回调

```java
Excel.read(inputStream)
        .fileName("users.xlsx")
        .sheet(0)
        .list(UserExcelDTO.class, user -> {
            // 每解析完成一行都会回调
            System.out.println(user.getName());
        })
        .parse();
```

### 批处理回调

适合导入数据量较大时分批落库。

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

### 读取图片

开启 `readPicture(true)` 后，模型字段可以使用 `ExcelReaderPictureModel`、`List<ExcelReaderPictureModel>` 或 `byte[]` 接收单元格图片。

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

### 自定义读取

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

### 读取 Map/CSV

`mapList()` 会把第一行作为表头，后续行转换为 `List<Map<String, Object>>`。

```java
List<Map<String, Object>> rows = Excel.read(inputStream)
        .fileName("users.csv")
        .mapList();
```

CSV 也可以按模型读取：

```java
List<UserExcelDTO> users = Excel.read(inputStream)
        .fileName("users.csv")
        .list(UserExcelDTO.class)
        .parse()
        .getList(UserExcelDTO.class);
```

## Large Data Export

大数据导出使用 `Excel.largeWrite(...)`。该模式会复用内部大数据写入器，适合百万级列表或需要按最大行数拆分 sheet 的场景。

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

同一个大数据导出文件可以追加多个列表；每次 `list(...)` 可指定模型和排除字段。

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

批量导出会并发生成多个 Excel 文件，并最终压缩为 zip。`ExcelBatchWriteBuilder(String outputDirPath, int maxPoolSize)` 可指定最大线程数。

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

导出 zip 到 HTTP 响应：

```java
new ExcelBatchWriteBuilder("/tmp/flux-excel", 4)
        .file("users")
        .sheet("用户")
        .list(users, UserExcelDTO.class)
        .export(request, response, "users");
```

## Event Read

事件模式读取适合超大文件解析，使用时需要提供行解析器和批处理器。

```java
Excel.eventRead(inputStream)
        .fileName("large-users.xlsx")
        .rowReader((sheetIndex, curRow, rowValue) -> {
            // 将原始行解析为业务对象
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

`rowReader` 接收到的 `sheetIndex` 从 `0` 开始，`curRow` 是 Excel 原始行号，`rowValue` 是当前行的字符串列表。事件模式适合大文件导入，但不会执行普通 `Excel.read(...)` 的模型注解映射、图片读取和模板校验。

## Compatible Writer API

推荐优先使用链式 API。需要兼容历史调用时，可以通过 `Excel.write(...).writer()` 获取底层 `ExcelWriter`。

```java
ExcelWriter writer = Excel.write(outputStream)
        .suffix(ExcelSuffixEnum.XLSX)
        .writer();

writer.setStreaming(true)
        .addModelList(users, "用户", true, UserExcelDTO.class)
        .addValidationOrComment("age", ageScope)
        .process(outputStream, "users.xlsx", ExcelSuffixEnum.XLSX);
```

常用兼容方法：

| 方法 | 说明 |
| --- | --- |
| `addModel(...)` | 添加单个模型 |
| `addModelList(...)` | 添加模型列表 |
| `addCustomColumn(...)` | 添加自定义单元格 |
| `addMergeCustomColumn(...)` | 添加合并单元格 |
| `addValidationOrComment(...)` | 添加批注、数字范围或下拉校验 |
| `setStreaming(boolean)` | 设置普通 `.xlsx` 是否使用 SXSSF |
| `selectSheet(String)` | 设置打开文件时选中的 sheet |
| `process(...)` / `export(...)` | 输出到流、文件或响应 |

## Spring Boot Starter

引入 starter 后，应用启动时会自动：

- 注册导出、导入模板目录的资源定位能力。
- 注册可选的导入模板排除器。
- 注册默认 `ExcelLargeListWriter` Bean。

DTO 元数据和模板内容不会在启动时扫描或解析，而是在导入、导出第一次使用对应 `Class` 或模板名时懒加载并缓存。

配置示例：

```yaml
excel:
  export-template-dir: excel-template
  import-template-dir: import-excel-template
```

资源目录示例：

```text
src/main/resources/
  excel-template/
    report-template.xlsx
  import-excel-template/
    users-import-template.xlsx
```

## Annotations

### `@ExcelWrite`

用于标记可导出的模型类。常用属性：

| 属性 | 说明 |
| --- | --- |
| `fillTitle` | 是否自动填充表头 |
| `freezeTitle` | 是否冻结标题 |
| `rowIndex` / `colIndex` | 默认起始行列 |
| `fillStyle` | 填充方向 |
| `titleStyleName` / `contentStyleName` | 标题和内容样式名 |
| `fillType` | 列表填充方式 |
| `themeName` / `customThemeName` | 主题配置 |
| `incrementSequenceNo` | 是否填充自增序号 |
| `filterTitle` | 是否开启表头筛选 |

### `@ExcelWriteProperty`

用于标记导出字段。常用属性：

| 属性 | 说明 |
| --- | --- |
| `titleName` | 表头名称 |
| `rowHeight` / `colWidth` | 行高和列宽 |
| `rowIndex` / `colIndex` | 字段固定行列 |
| `formatPattern` | 输出格式 |
| `formatter` | 自定义导出格式化器 |
| `commentText` | 批注内容 |
| `dropDownOptions` | 下拉选项 |
| `disable` | 是否禁用该字段 |
| `index` | 排序号 |

### `@ExcelRead`

用于标记可导入的模型类。常用属性：

| 属性 | 说明 |
| --- | --- |
| `checkTitle` | 是否校验表头 |
| `enableSeparator` | 是否启用分隔符解析 |

### `@ExcelReadProperty`

用于标记导入字段。常用属性：

| 属性 | 说明 |
| --- | --- |
| `titleName` | 对应表头名称 |
| `separator` | 分隔符 |
| `formatPattern` | 读取格式 |
| `formatter` | 自定义读取格式化器 |
| `checkNull` | 是否校验空值 |
| `disable` | 是否禁用该字段 |

## Custom Formatters

### 导出格式化器

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

### 导入格式化器

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

非 Spring Boot 项目默认也会按传入的模型 `Class` 懒加载元数据。模板可以通过 `ExcelMetadataRegistry` 配置目录后按模板名懒加载，也可以继续使用旧 API 主动预热：

```java
ExcelMetadataRegistry.configureTemplateDirectory("/path/to/export/templates", "/path/to/import/templates", null);

// 可选预热，非必须
ExcelMetadataRegistry.loadExcelTemplate("/path/to/export/templates");
ExcelMetadataRegistry.loadImportExcelTemplate("/path/to/import/templates", null);
```

如需显式热加载，可清理缓存，下一次导入导出会重新解析：

```java
ExcelMetadataRegistry.clearModel(UserExcelDTO.class);
ExcelMetadataRegistry.clearTemplate("user-template.xlsx");
```

`ExcelBootLoader` 仍作为兼容门面保留，历史调用会委托到 `ExcelMetadataRegistry`。

## Build And Test

编译：

```bash
mvn -pl flux-excel-core -am compile
```

运行测试：

```bash
mvn -pl flux-excel-core test
```

完整构建：

```bash
mvn clean install
```

## Project Structure

```text
flux-excel/
  flux-excel-core/                 # 核心导入导出能力
  flux-excel-spring-boot-starter/  # Spring Boot 自动配置
  pom.xml
  README.MD
  LICENSE
```

## Roadmap

- 继续拆分 `BaseExcelWriter`，降低核心写入逻辑复杂度。
- 完善单元测试和大数据导出性能测试。
- 补充更多 Spring Boot 示例项目。

## Contributing

欢迎提交 Issue 和 Pull Request。建议在提交前执行：

```bash
mvn clean test
```

提交 PR 时请说明：

- 改动目的和使用场景。
- 是否包含兼容性影响。
- 是否新增或更新测试。

## License

flux-excel 使用 [Apache License 2.0](LICENSE) 开源协议。
