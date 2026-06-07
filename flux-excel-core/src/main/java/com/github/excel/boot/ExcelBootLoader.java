package com.github.excel.boot;

import com.github.excel.model.ExcelCacheImportModel;
import com.github.excel.model.ExcelCacheModel;
import com.github.excel.model.ExcelExpressionModel;
import com.github.excel.model.ExcelImportTemplateCacheModel;
import com.github.excel.model.ExcelTemplateTitleModel;
import com.github.excel.read.facade.AbstractReaderTemplateExclude;

import java.util.List;
import java.util.Map;

/**
 * @Author: Vachel Wang
 * @Date: 2026/4/24
 * @Description: Excel 引导加载器兼容门面，推荐使用 {@link ExcelMetadataRegistry}
 */
public class ExcelBootLoader {

    public interface TemplateResourceResolver extends ExcelMetadataRegistry.TemplateResourceResolver {
    }

    public static void configureTemplateDirectory(String exportTemplateDir, String importTemplateDir, AbstractReaderTemplateExclude excluder) {
        ExcelMetadataRegistry.configureTemplateDirectory(exportTemplateDir, importTemplateDir, excluder);
    }

    public static void configureTemplateResolver(ExcelMetadataRegistry.TemplateResourceResolver templateResourceResolver) {
        ExcelMetadataRegistry.configureTemplateResolver(templateResourceResolver);
    }

    public static String getTemplateDirPath(){
        return ExcelMetadataRegistry.getTemplateDirPath();
    }

    public static ExcelCacheModel getExcelCacheMapValue(Class key){
        return ExcelMetadataRegistry.getExcelCacheMapValue(key);
    }

    public static Map<String, List<ExcelExpressionModel>> getExcelTemplateCacheMapValue(String key){
        return ExcelMetadataRegistry.getExcelTemplateCacheMapValue(key);
    }

    public static ExcelCacheImportModel getExcelCacheImportMapValue(Class key){
        return ExcelMetadataRegistry.getExcelCacheImportMapValue(key);
    }

    public static Map<Integer, List<ExcelTemplateTitleModel>> getExcelTemplateTitleCacheMapValue(String key){
        return ExcelMetadataRegistry.getExcelTemplateTitleCacheMapValue(key);
    }

    public static byte[] getExcelImportTemplateFileCacheMapValue(String key){
        return ExcelMetadataRegistry.getExcelImportTemplateFileCacheMapValue(key);
    }

    public static Map<String,List<ExcelImportTemplateCacheModel>> getExcelImportTemplateCacheMapValue(String key){
        return ExcelMetadataRegistry.getExcelImportTemplateCacheMapValue(key);
    }

    public static void clearExportModel(Class key) {
        ExcelMetadataRegistry.clearExportModel(key);
    }

    public static void clearImportModel(Class key) {
        ExcelMetadataRegistry.clearImportModel(key);
    }

    public static void clearModel(Class key) {
        ExcelMetadataRegistry.clearModel(key);
    }

    public static void clearExportTemplate(String templateName) {
        ExcelMetadataRegistry.clearExportTemplate(templateName);
    }

    public static void clearImportTemplate(String templateName) {
        ExcelMetadataRegistry.clearImportTemplate(templateName);
    }

    public static void clearTemplate(String templateName) {
        ExcelMetadataRegistry.clearTemplate(templateName);
    }

    public static void clearAll() {
        ExcelMetadataRegistry.clearAll();
    }

    public static void loadModel(String ... packagePathArray) {
        ExcelMetadataRegistry.loadModel(packagePathArray);
    }

    public static void loadExcelTemplate(String templatePath) {
        ExcelMetadataRegistry.loadExcelTemplate(templatePath);
    }

    public static void loadImportExcelTemplate(String templatePath, AbstractReaderTemplateExclude excluder) {
        ExcelMetadataRegistry.loadImportExcelTemplate(templatePath, excluder);
    }

    public static void loadExcelTemplate(List<Map<String, Object>> fileList) {
        ExcelMetadataRegistry.loadExcelTemplate(fileList);
    }

    public static void loadImportExcelTemplate(List<Map<String, Object>> fileList, AbstractReaderTemplateExclude excluder) {
        ExcelMetadataRegistry.loadImportExcelTemplate(fileList, excluder);
    }
}
