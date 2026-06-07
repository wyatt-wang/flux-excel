package com.github.boot;

import com.github.excel.annotation.ExcelRead;
import com.github.excel.annotation.ExcelWrite;
import com.github.excel.boot.ExcelBootLoader;
import com.github.excel.boot.ExcelMetadataRegistry;
import com.github.excel.exception.ExcelReaderException;
import com.github.excel.exception.ExcelWriterException;
import com.github.excel.model.ExcelBaseModel;
import com.github.excel.model.ExcelCacheImportModel;
import com.github.excel.model.ExcelCacheModel;
import com.github.model.UserExcelDto;
import com.github.model.UserExcelDtoImportBean;
import com.github.read.template.ImportTemplateExcluder;
import org.junit.After;
import org.junit.Before;
import org.junit.Test;

import java.net.URL;
import java.util.List;
import java.util.Map;

import static org.junit.Assert.*;

public class ExcelBootLoaderLazyLoadTest {

    @Before
    public void setUp() {
        ExcelMetadataRegistry.clearAll();
    }

    @After
    public void tearDown() {
        ExcelMetadataRegistry.clearAll();
    }

    @Test
    public void getExcelCacheMapValueParsesExportModelWithoutPackageScan() {
        ExcelCacheModel cacheModel = ExcelMetadataRegistry.getExcelCacheMapValue(UserExcelDto.class);

        assertNotNull(cacheModel);
        assertNotNull(cacheModel.getExcelWrite());
        assertFalse(cacheModel.getFieldModelList().isEmpty());
        assertTrue(cacheModel.getFieldModelMap().containsKey("姓名"));
    }

    @Test
    public void getExcelCacheImportMapValueParsesImportModelWithoutPackageScan() {
        ExcelCacheImportModel cacheModel = ExcelMetadataRegistry.getExcelCacheImportMapValue(UserExcelDtoImportBean.class);

        assertNotNull(cacheModel);
        assertNotNull(cacheModel.getExcelRead());
        assertNotNull(cacheModel.getValidation());
        assertTrue(cacheModel.getFieldModelMap().containsKey("姓名"));
    }

    @Test(expected = ExcelWriterException.class)
    public void getExcelCacheMapValueRejectsClassWithoutExcelWrite() {
        ExcelMetadataRegistry.getExcelCacheMapValue(ReadOnlyModel.class);
    }

    @Test(expected = ExcelReaderException.class)
    public void getExcelCacheImportMapValueRejectsClassWithoutExcelRead() {
        ExcelMetadataRegistry.getExcelCacheImportMapValue(WriteOnlyModel.class);
    }

    @Test
    public void getExcelTemplateCacheMapValueParsesExportTemplateLazilyFromConfiguredDirectory() {
        URL templateDir = ExcelBootLoader.class.getClassLoader().getResource("excel-template");
        assertNotNull(templateDir);
        ExcelMetadataRegistry.configureTemplateDirectory(templateDir.getPath(), null, null);

        Map<String, List<com.github.excel.model.ExcelExpressionModel>> expressionMap =
                ExcelMetadataRegistry.getExcelTemplateCacheMapValue("test.xlsx");
        Map<Integer, List<com.github.excel.model.ExcelTemplateTitleModel>> titleMap =
                ExcelMetadataRegistry.getExcelTemplateTitleCacheMapValue("test.xlsx");

        assertNotNull(expressionMap);
        assertNotNull(titleMap);
        assertFalse(expressionMap.isEmpty());
        assertFalse(titleMap.isEmpty());

        ExcelMetadataRegistry.clearExportTemplate("test.xlsx");
        assertNotSame(expressionMap, ExcelMetadataRegistry.getExcelTemplateCacheMapValue("test.xlsx"));
    }

    @Test
    public void getExcelImportTemplateCacheMapValueParsesImportTemplateLazilyFromConfiguredDirectory() {
        URL templateDir = ExcelBootLoader.class.getClassLoader().getResource("import-excel-template");
        assertNotNull(templateDir);
        ImportTemplateExcluder excluder = new ImportTemplateExcluder();
        excluder.addTemplateExclude();
        ExcelMetadataRegistry.configureTemplateDirectory(null, templateDir.getPath(), excluder);

        Map<String, List<com.github.excel.model.ExcelImportTemplateCacheModel>> templateMap =
                ExcelMetadataRegistry.getExcelImportTemplateCacheMapValue("project-bids.xlsx");
        byte[] templateBytes = ExcelMetadataRegistry.getExcelImportTemplateFileCacheMapValue("project-bids.xlsx");

        assertNotNull(templateMap);
        assertFalse(templateMap.isEmpty());
        assertNotNull(templateBytes);
        assertTrue(templateBytes.length > 0);

        ExcelMetadataRegistry.clearImportTemplate("project-bids.xlsx");
        assertNotSame(templateMap, ExcelMetadataRegistry.getExcelImportTemplateCacheMapValue("project-bids.xlsx"));
    }

    @Test
    public void excelBootLoaderDelegatesToMetadataRegistryForBackwardCompatibility() {
        ExcelCacheModel registryModel = ExcelMetadataRegistry.getExcelCacheMapValue(UserExcelDto.class);
        ExcelCacheModel bootLoaderModel = ExcelBootLoader.getExcelCacheMapValue(UserExcelDto.class);

        assertSame(registryModel, bootLoaderModel);
    }

    @ExcelRead
    public static class ReadOnlyModel extends ExcelBaseModel {
    }

    @ExcelWrite
    public static class WriteOnlyModel extends ExcelBaseModel {
    }
}
