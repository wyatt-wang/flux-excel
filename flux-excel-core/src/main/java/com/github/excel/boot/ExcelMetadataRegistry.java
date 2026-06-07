package com.github.excel.boot;

import com.github.excel.annotation.*;
import com.github.excel.constant.ExcelConstant;
import com.github.excel.exception.ExcelReaderException;
import com.github.excel.exception.ExcelWriterException;
import com.github.excel.model.*;
import com.github.excel.read.facade.AbstractReaderTemplateExclude;
import com.github.excel.util.PackageUtil;
import com.github.excel.util.StringUtil;
import com.github.excel.write.ExcelTheme;
import com.github.excel.write.ExcelThemeRegistry;
import com.google.common.base.Throwables;
import com.google.common.collect.Lists;
import com.google.common.collect.Maps;
import lombok.extern.slf4j.Slf4j;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.time.Clock;
import java.util.*;
import java.util.stream.Collectors;

/**
 * @Author: Vachel Wang
 * @Date: 2026/4/24
 * @Description: Excel 元数据注册中心
 */
@Slf4j
public class ExcelMetadataRegistry {
    /**
     * 导出缓存
     */
    private static final Map<Class, ExcelCacheModel> EXCEL_CACHE_MAP = Maps.newConcurrentMap();
    /**
     * 导入缓存
     */
    private static final Map<Class, ExcelCacheImportModel> EXCEL_CACHE_IMPORT_MAP = Maps.newConcurrentMap();

    /**
     * 模板表达式缓存
     */
    private static final Map<String, Map<String, List<ExcelExpressionModel>>> EXCEL_TEMPLATE_CACHE_MAP = Maps.newConcurrentMap();

    /**
     * 模板标题缓存
     */
    private static final Map<String, Map<Integer, List<ExcelTemplateTitleModel>>> EXCEL_TEMPLATE_TITLE_CACHE_MAP = Maps.newConcurrentMap();

    /**
     * 导入模板缓存
     */
    private static final Map<String, byte[]> EXCEL_IMPORT_TEMPLATE_FILE_CACHE_MAP = Maps.newConcurrentMap();

    /**
     * 导入模板缓存 key->文件名, value-> (key->sheet名,value->模板内容)
     */
    private static final Map<String, Map<String,List<ExcelImportTemplateCacheModel>>> EXCEL_IMPORT_TEMPLATE_CACHE_MAP = Maps.newConcurrentMap();

    /**
     * 模板文件夹地址
     */
    private static String TEMPLATE_DIR_PATH ;
    /**
     * 导出模板文件夹地址
     */
    private static String EXPORT_TEMPLATE_DIR_PATH;
    /**
     * 导入模板文件夹地址
     */
    private static String IMPORT_TEMPLATE_DIR_PATH;
    /**
     * 导入模板排除器
     */
    private static AbstractReaderTemplateExclude IMPORT_TEMPLATE_EXCLUDER;
    /**
     * Spring Boot 等外部环境注册的模板资源定位器
     */
    private static TemplateResourceResolver TEMPLATE_RESOURCE_RESOLVER;

    private static final Object EXPORT_TEMPLATE_LOAD_LOCK = new Object();
    private static final Object IMPORT_TEMPLATE_LOAD_LOCK = new Object();

    public interface TemplateResourceResolver {
        List<Map<String, Object>> resolve(String templateDir, String templateName);
    }

    public static void configureTemplateDirectory(String exportTemplateDir, String importTemplateDir, AbstractReaderTemplateExclude excluder) {
        EXPORT_TEMPLATE_DIR_PATH = exportTemplateDir;
        IMPORT_TEMPLATE_DIR_PATH = importTemplateDir;
        IMPORT_TEMPLATE_EXCLUDER = excluder;
    }

    public static void configureTemplateResolver(TemplateResourceResolver templateResourceResolver) {
        TEMPLATE_RESOURCE_RESOLVER = templateResourceResolver;
    }

    /**
     * 获取导出模板文件夹名称
     * @return
     */
    public static String getTemplateDirPath(){
        return TEMPLATE_DIR_PATH;
    }

    /**
     * 获取导出缓存
     * @param key 模板名称
     * @return
     */
    public static ExcelCacheModel getExcelCacheMapValue(Class key){
        if (Objects.isNull(key)) {
            return null;
        }
        validateModelClass(key, ExcelWriterException::new);
        ExcelWrite excelWrite = (ExcelWrite) key.getAnnotation(ExcelWrite.class);
        if (Objects.isNull(excelWrite)) {
            throw new ExcelWriterException("Model class missing @ExcelWrite: " + key.getName());
        }
        return EXCEL_CACHE_MAP.computeIfAbsent(key, cla -> parseExportModel(cla, excelWrite));
    }

    /**
     * 获取导出模板表达式缓存
     * @param key 模板名称
     * @return
     */
    public static Map<String, List<ExcelExpressionModel>> getExcelTemplateCacheMapValue(String key){
        loadExportTemplateIfNecessary(key);
        return EXCEL_TEMPLATE_CACHE_MAP.get(key);
    }

    /**
     * 获取导入模板缓存
     * @param key 模型class
     * @return
     */
    public static ExcelCacheImportModel getExcelCacheImportMapValue(Class key){
        if (Objects.isNull(key)) {
            return null;
        }
        validateModelClass(key, ExcelReaderException::new);
        ExcelRead excelRead = (ExcelRead) key.getAnnotation(ExcelRead.class);
        if (Objects.isNull(excelRead)) {
            throw new ExcelReaderException("Model class missing @ExcelRead: " + key.getName());
        }
        return EXCEL_CACHE_IMPORT_MAP.computeIfAbsent(key, cla -> parseImportModel(cla, excelRead));
    }

    /**
     * 获取导出标题缓存
     * @param key 模板名称
     * @return
     */
    public static Map<Integer, List<ExcelTemplateTitleModel>> getExcelTemplateTitleCacheMapValue(String key){
        loadExportTemplateIfNecessary(key);
        return EXCEL_TEMPLATE_TITLE_CACHE_MAP.get(key);
    }

    /**
     * 获取导入模板文件流
     * @param key 导入模板文件名
     * @return
     */
    public static byte[] getExcelImportTemplateFileCacheMapValue(String key){
        loadImportTemplateIfNecessary(key);
        return EXCEL_IMPORT_TEMPLATE_FILE_CACHE_MAP.get(key);
    }

    /**
     * 获取导入模板缓存
     * @param key 模板文件名称
     * @return
     */
    public static Map<String,List<ExcelImportTemplateCacheModel>> getExcelImportTemplateCacheMapValue(String key){
        loadImportTemplateIfNecessary(key);
        return EXCEL_IMPORT_TEMPLATE_CACHE_MAP.get(key);
    }

    public static void clearExportModel(Class key) {
        if (Objects.nonNull(key)) {
            EXCEL_CACHE_MAP.remove(key);
        }
    }

    public static void clearImportModel(Class key) {
        if (Objects.nonNull(key)) {
            EXCEL_CACHE_IMPORT_MAP.remove(key);
        }
    }

    public static void clearModel(Class key) {
        clearExportModel(key);
        clearImportModel(key);
    }

    public static void clearExportTemplate(String templateName) {
        if (StringUtil.notEmpty(templateName)) {
            EXCEL_TEMPLATE_CACHE_MAP.remove(templateName);
            EXCEL_TEMPLATE_TITLE_CACHE_MAP.remove(templateName);
        }
    }

    public static void clearImportTemplate(String templateName) {
        if (StringUtil.notEmpty(templateName)) {
            EXCEL_IMPORT_TEMPLATE_FILE_CACHE_MAP.remove(templateName);
            EXCEL_IMPORT_TEMPLATE_CACHE_MAP.remove(templateName);
        }
    }

    public static void clearTemplate(String templateName) {
        clearExportTemplate(templateName);
        clearImportTemplate(templateName);
    }

    public static void clearAll() {
        EXCEL_CACHE_MAP.clear();
        EXCEL_CACHE_IMPORT_MAP.clear();
        EXCEL_TEMPLATE_CACHE_MAP.clear();
        EXCEL_TEMPLATE_TITLE_CACHE_MAP.clear();
        EXCEL_IMPORT_TEMPLATE_FILE_CACHE_MAP.clear();
        EXCEL_IMPORT_TEMPLATE_CACHE_MAP.clear();
    }


    /**
     * 加载并缓存excel模型信息
     *
     * @param packagePathArray
     */
    @SuppressWarnings("unchecked")
    public static void loadModel(String ... packagePathArray) {
        if (Objects.isNull(packagePathArray) || packagePathArray.length == ExcelConstant.ZERO_SHORT) {
            throw new IllegalArgumentException("packagePath can't be null");
        }
        Clock loadClock = Clock.systemUTC();
        long startMills = loadClock.millis();
        log.info("flux-excel loading model");
        List<Class<?>> classList = new ArrayList<>();
        for(String packagePath:packagePathArray) {
            List<Class<?>> clasList = PackageUtil.getClass(packagePath, true);
            classList.addAll(clasList);
        }
        for (Class cla : classList) {
            ExcelWrite excelWrite = (ExcelWrite) cla.getAnnotation(ExcelWrite.class);
            ExcelRead excelRead = (ExcelRead) cla.getAnnotation(ExcelRead.class);
            if ((Objects.isNull(excelWrite) && Objects.isNull(excelRead)) || !ExcelBaseModel.class.isAssignableFrom(cla)) {
                continue;
            }
            if(Objects.nonNull(excelWrite)) {
                EXCEL_CACHE_MAP.put(cla, parseExportModel(cla, excelWrite));
            }
            if (Objects.nonNull(excelRead)){
                EXCEL_CACHE_IMPORT_MAP.put(cla, parseImportModel(cla, excelRead));
            }
        }
        long loadMillis = loadClock.millis() - startMills;
        log.info("flux-excel load success by {} millisecond", loadMillis);
    }

    /**
     * @Author: Vachel Wang
     * @Date: 2026/4/24
     * @Description: 加载导出模型
     */
    private static ExcelCacheModel parseExportModel(Class cla, ExcelWrite excelWrite) {
        ExcelCacheModel cacheModel = new ExcelCacheModel();
        cacheModel.setExcelWrite(excelWrite);
        List<ExcelCacheFieldModel> fieldModelList = Lists.newArrayList();
        Map<String,ExcelCacheFieldModel> fieldModelMap = Maps.newHashMap();
        List<Field> excelFields = getModelFields(cla);
        List<Short> listHeight = new ArrayList<>(excelFields.size());
        setIncrementSequenceNoField(excelWrite,fieldModelList,fieldModelMap);
        ExcelTheme theme = getTheme(excelWrite);
        for (Field excelField : excelFields) {
            ExcelWriteProperty exportCell = excelField.getDeclaredAnnotation(ExcelWriteProperty.class);
            if (Objects.isNull(exportCell) || exportCell.disable()) {
                continue;
            }
            ExcelCacheFieldModel fieldModel = new ExcelCacheFieldModel();

            short titleRowHeight = resolveRowHeight(exportCell);
            short contentRowHeight = resolveRowHeight(exportCell);
            short colWidth = resolveColWidth(exportCell);
            if (theme != null && theme != ExcelThemeRegistry.getTheme(null)) {
                if (ExcelConstant.MINUS_TWO_SHORT == contentRowHeight) {
                    contentRowHeight = theme.getContentRowHeight();
                    titleRowHeight = theme.getTitleRowHeight();
                }
                if (ExcelConstant.MINUS_TWO_SHORT == colWidth) {
                    colWidth = theme.getColWidth();
                }
            }
            listHeight.add(contentRowHeight);
            //当属性是Map类型的时候  设置个标记
            if(excelField.getType().equals(Map.class)){
                fieldModel.setMap(true);
            }else {
                fieldModel.setMap(false);
            }
            String methodName = StringUtil.concatMethodName(excelField.getName(),ExcelConstant.GET_STR);
            try {
                Method getMethod = cla.getMethod(methodName);
                boolean isDate = false;
                if (Date.class.isAssignableFrom(getMethod.getReturnType()) || Calendar.class.isAssignableFrom(getMethod.getReturnType())) {
                    isDate = true;
                }
                String titleStyleName = StringUtil.notEmpty(exportCell.titleStyleName()) ? exportCell.titleStyleName() : StringUtil.notEmpty(excelWrite.titleStyleName()) ? excelWrite.titleStyleName() : theme.getTitleRowStyleName();
                String contentStyleName = StringUtil.notEmpty(exportCell.contentStyleName()) ? exportCell.contentStyleName() : StringUtil.notEmpty(excelWrite.contentStyleName()) ? excelWrite.contentStyleName() : isDate ? theme.getOddRowStyleDateName(): theme.getOddRowStyleName();
                String evenRowStyleName = StringUtil.notEmpty(exportCell.contentStyleName()) ? exportCell.contentStyleName() : StringUtil.notEmpty(excelWrite.contentStyleName()) ? excelWrite.contentStyleName() : isDate ? theme.getEvenRowStyleDateName(): theme.getEvenRowStyleName();
                fieldModel.setExportCell(exportCell);
                fieldModel.setGetMethod(getMethod);
                fieldModel.setLinkNameGetMethod(getLinkNameGetMethod(cla, exportCell));
                fieldModel.setFieldName(excelField.getName());
                fieldModel.setTitleStyleName(titleStyleName);
                fieldModel.setContentStyleName(contentStyleName);
                fieldModel.setEvenRowStyleName(evenRowStyleName);
                fieldModel.setTitleRowHeight(titleRowHeight);
                fieldModel.setContentRowHeight(contentRowHeight);
                fieldModel.setColWidth(colWidth);
                fieldModel.setIndex(exportCell.index());
                List<String> headNames = resolveWriteHeadNames(exportCell);
                fieldModel.setHeadNames(headNames);
                fieldModel.setTitleName(headNames.get(headNames.size() - 1));
                fieldModelList.add(fieldModel);
                fieldModelMap.put(exportCell.titleName(), fieldModel);
                fieldModelMap.put(canonicalHeadKey(headNames), fieldModel);
            } catch (NoSuchMethodException e) {
                log.error(Throwables.getStackTraceAsString(e));
                throw new ExcelWriterException("load export model failed ,cause:" + e.getMessage());
            }
        }
        if(!fieldModelList.isEmpty()){
            // 排序
            fieldModelList = fieldModelList.stream().sorted(Comparator.comparing(ExcelCacheFieldModel::getIndex)).collect(Collectors.toList());
            short listRowHeight = fieldModelList.stream().map(ExcelCacheFieldModel::getContentRowHeight).max(Comparator.comparing(Short::shortValue)).orElse(ExcelConstant.MINUS_TWO_SHORT);
            for (ExcelCacheFieldModel excelCacheFieldModel : fieldModelList) {
                excelCacheFieldModel.setListRowHeight(listRowHeight);
            }
        }
        cacheModel.setFieldModelList(fieldModelList);
        cacheModel.setFieldModelMap(fieldModelMap);
        cacheModel.setMaxHeadDepth(fieldModelList.stream()
                .map(ExcelCacheFieldModel::getHeadNames)
                .filter(Objects::nonNull)
                .map(List::size)
                .max(Integer::compareTo)
                .orElse(ExcelConstant.ONE_INT));
        return cacheModel;
    }

    private static void setIncrementSequenceNoField(ExcelWrite excelWrite, List<ExcelCacheFieldModel> fieldModelList, Map<String, ExcelCacheFieldModel> fieldModelMap){
        if (excelWrite.incrementSequenceNo()) {
            ExcelTheme theme = getTheme(excelWrite);

            String titleStyleName =  StringUtil.notEmpty(excelWrite.titleStyleName()) ? excelWrite.titleStyleName() : theme.getTitleRowStyleName();
            String contentStyleName = StringUtil.notEmpty(excelWrite.contentStyleName()) ? excelWrite.contentStyleName() : theme.getOddRowStyleName();
            String evenRowStyleName = StringUtil.notEmpty(excelWrite.contentStyleName()) ? excelWrite.contentStyleName() : theme.getEvenRowStyleName();

            ExcelCacheFieldModel fieldModel = new ExcelCacheFieldModel();
            fieldModel.setMap(false);
            fieldModel.setFieldName(ExcelConstant.INCREMENT_SEQUENCE_NO_FIELD_NAME);
            fieldModel.setTitleStyleName(titleStyleName);
            fieldModel.setContentStyleName(contentStyleName);
            fieldModel.setEvenRowStyleName(evenRowStyleName);
            fieldModel.setColWidth(ExcelConstant.SHORT_100);
            fieldModel.setTitleRowHeight(ExcelConstant.MINUS_TWO_SHORT);
            fieldModel.setContentRowHeight(ExcelConstant.MINUS_TWO_SHORT);
            fieldModel.setIndex(ExcelConstant.INT_NEGATIVE_9999);
            fieldModel.setTitleName(excelWrite.incrementSequenceTitle());
            fieldModel.setHeadNames(Collections.singletonList(excelWrite.incrementSequenceTitle()));
            fieldModelList.add(fieldModel);
            fieldModelMap.put(ExcelConstant.INCREMENT_SEQUENCE_NO_FIELD_NAME, fieldModel);
        }
    }

    private static ExcelTheme getTheme(ExcelWrite excelWrite) {
        if (Objects.isNull(excelWrite)) {
            return ExcelThemeRegistry.getTheme((String) null);
        }
        if (StringUtil.notEmpty(excelWrite.customThemeName()) && !"NONE".equalsIgnoreCase(excelWrite.customThemeName())) {
            return ExcelThemeRegistry.getTheme(excelWrite.customThemeName());
        }
        return ExcelThemeRegistry.getTheme(excelWrite.themeName().name());
    }

    private static List<Field> getModelFields(Class cla) {
        LinkedHashMap<String, Field> fieldMap = Maps.newLinkedHashMap();
        List<Class> hierarchy = Lists.newArrayList();
        Class current = cla;
        while (Objects.nonNull(current) && current != Object.class) {
            hierarchy.add(current);
            current = current.getSuperclass();
        }
        Collections.reverse(hierarchy);
        for (Class currentClass : hierarchy) {
            for (Field field : currentClass.getDeclaredFields()) {
                fieldMap.put(field.getName(), field);
            }
        }
        return Lists.newArrayList(fieldMap.values());
    }

    private static short resolveRowHeight(ExcelWriteProperty exportCell) {
        return exportCell.rowHeightPoints() != ExcelConstant.MINUS_TWO_SHORT ? exportCell.rowHeightPoints() : exportCell.rowHeight();
    }

    private static short resolveColWidth(ExcelWriteProperty exportCell) {
        if (exportCell.colWidthChars() == ExcelConstant.MINUS_TWO_SHORT) {
            return exportCell.colWidth();
        }
        if (exportCell.colWidthChars() == ExcelConstant.MINUS_ONE_SHORT) {
            return ExcelConstant.MINUS_ONE_SHORT;
        }
        return (short) Math.round(exportCell.colWidthChars() * 256 / ExcelConstant.PIXEL_RATE);
    }

    private static Method getLinkNameGetMethod(Class cla, ExcelWriteProperty exportCell) throws NoSuchMethodException {
        if (StringUtil.isEmpty(exportCell.linkNameField())) {
            return null;
        }
        return cla.getMethod(StringUtil.concatMethodName(exportCell.linkNameField(), ExcelConstant.GET_STR));
    }

    /**
     * 加载导入模型
     * @param cla
     * @param excelRead
     */
    private static ExcelCacheImportModel parseImportModel(Class cla, ExcelRead excelRead) {
        ExcelCacheImportModel cacheModel = new ExcelCacheImportModel();
        cacheModel.setExcelRead(excelRead);
        cacheModel.setValidation((ExcelValidation)cla.getAnnotation(ExcelValidation.class));
        Map<String,ExcelCacheImportModel.ExcelCacheImportFieldModel> fieldModelMap = Maps.newHashMap();
        for (Field excelField : getModelFields(cla)) {
            ExcelReadProperty importProperty = excelField.getDeclaredAnnotation(ExcelReadProperty.class);
            if (Objects.isNull(importProperty) || importProperty.disable()) {
                continue;
            }
            ExcelCacheImportModel.ExcelCacheImportFieldModel fieldModel = new ExcelCacheImportModel.ExcelCacheImportFieldModel();
            String methodName = StringUtil.concatMethodName(excelField.getName(),ExcelConstant.SET_STR);
            try {
                Method setMethod = cla.getMethod(methodName, excelField.getType());
                fieldModel.setImportProperty(importProperty);
                fieldModel.setSetMethod(setMethod);
                fieldModel.setField(excelField.getName());
                List<String> headNames = resolveReadHeadNames(importProperty);
                fieldModel.setHeadNames(headNames);
                String modelMapKey = excelRead.enableSeparator() ? importProperty.titleName() + importProperty.separator() : importProperty.titleName();
                fieldModelMap.put(modelMapKey, fieldModel);
                fieldModelMap.put(canonicalHeadKey(headNames), fieldModel);
                if (excelRead.enableSeparator()) {
                    fieldModelMap.put(canonicalHeadKey(Collections.singletonList(modelMapKey)), fieldModel);
                }
            } catch (NoSuchMethodException e) {
                log.error(Throwables.getStackTraceAsString(e));
                throw new ExcelReaderException("load import model failed ,cause:" + e.getMessage());
            }
        }
        cacheModel.setFieldModelMap(fieldModelMap);
        cacheModel.setMaxHeadDepth(fieldModelMap.values().stream()
                .map(ExcelCacheImportModel.ExcelCacheImportFieldModel::getHeadNames)
                .filter(Objects::nonNull)
                .map(List::size)
                .max(Integer::compareTo)
                .orElse(ExcelConstant.ONE_INT));
        return cacheModel;
    }

    private static List<String> resolveWriteHeadNames(ExcelWriteProperty exportCell) {
        if (exportCell.head().length > ExcelConstant.ZERO_SHORT) {
            return normalizeHeadNames(Arrays.asList(exportCell.head()), exportCell.titleName());
        }
        if (exportCell.value().length > ExcelConstant.ZERO_SHORT) {
            return normalizeHeadNames(Arrays.asList(exportCell.value()), exportCell.titleName());
        }
        return normalizeHeadNames(Collections.singletonList(exportCell.titleName()), exportCell.titleName());
    }

    private static List<String> resolveReadHeadNames(ExcelReadProperty importProperty) {
        if (importProperty.head().length > ExcelConstant.ZERO_SHORT) {
            return normalizeHeadNames(Arrays.asList(importProperty.head()), importProperty.titleName());
        }
        if (importProperty.value().length > ExcelConstant.ZERO_SHORT) {
            return normalizeHeadNames(Arrays.asList(importProperty.value()), importProperty.titleName());
        }
        return normalizeHeadNames(Collections.singletonList(importProperty.titleName()), importProperty.titleName());
    }

    private static List<String> normalizeHeadNames(List<String> source, String fallbackTitle) {
        List<String> headNames = source.stream()
                .filter(StringUtil::notEmpty)
                .map(String::trim)
                .collect(Collectors.toList());
        if (headNames.isEmpty()) {
            headNames.add(fallbackTitle);
        }
        return headNames;
    }

    private static String canonicalHeadKey(List<String> headNames) {
        if (headNames == null || headNames.isEmpty()) {
            return ExcelConstant.NULL_STR;
        }
        List<String> normalized = new ArrayList<>();
        String previous = null;
        for (String headName : headNames) {
            if (StringUtil.isEmpty(headName)) {
                continue;
            }
            String current = headName.trim();
            if (!current.equals(previous)) {
                normalized.add(current);
            }
            previous = current;
        }
        return String.join("\u001F", normalized);
    }


    /**
     *  解析模板并缓存
     */
    public static void loadExcelTemplate(String templatePath) {
        TEMPLATE_DIR_PATH = templatePath+=ExcelConstant.FILE_SEPARATOR;
        EXPORT_TEMPLATE_DIR_PATH = templatePath;
        Clock loadClock = Clock.systemUTC();
        long startMills = loadClock.millis();
        log.info("flux-excel loading export template");
        try {
            //读所有的文件
            File directory = new File(templatePath);
            if (directory.isDirectory()) {
                File[] files = Optional.ofNullable(directory.listFiles()).orElse(new File[ExcelConstant.ZERO_SHORT]);
                List<File> fileList = getFileList(files);
                ExcelTemplateLoader.loadExportTemplate(fileList, EXCEL_TEMPLATE_CACHE_MAP, EXCEL_TEMPLATE_TITLE_CACHE_MAP);
            }else {
                log.error("directory.is.empty");
                throw new ExcelReaderException("directory.is.empty");
            }
        }catch (Exception e){
            log.error(Throwables.getStackTraceAsString(e));
            throw new ExcelReaderException("template.load.fail");
        }
        long loadMillis = loadClock.millis() - startMills;
        log.info("flux-excel load template success by {} millisecond", loadMillis);
    }

    /**
     *  导入模板解析并缓存
     */
    public static void loadImportExcelTemplate(String templatePath, AbstractReaderTemplateExclude excluder) {
        TEMPLATE_DIR_PATH = templatePath+=ExcelConstant.FILE_SEPARATOR;
        IMPORT_TEMPLATE_DIR_PATH = templatePath;
        IMPORT_TEMPLATE_EXCLUDER = excluder;
        Clock loadClock = Clock.systemUTC();
        long startMills = loadClock.millis();
        log.info("flux-excel loading import template");
        try {
            //读所有的文件
            File directory = new File(templatePath);
            if (directory.isDirectory()) {
                File[] files = Optional.ofNullable(directory.listFiles()).orElse(new File[ExcelConstant.ZERO_SHORT]);
                List<File> fileList = getFileList(files);
                ExcelTemplateLoader.loadImportTemplate(fileList,EXCEL_IMPORT_TEMPLATE_FILE_CACHE_MAP,EXCEL_IMPORT_TEMPLATE_CACHE_MAP,excluder);
            }else {
                log.error("directory.is.empty");
                throw new ExcelReaderException("directory.is.empty");
            }
        }catch (Exception e){
            log.error(Throwables.getStackTraceAsString(e));
            throw new ExcelReaderException("template.load.fail");
        }
        long loadMillis = loadClock.millis() - startMills;
        log.info("flux-excel load template success by {} millisecond", loadMillis);
    }

    /**
     *  解析模板并缓存
     */
    public static void loadExcelTemplate(List<Map<String, Object>> fileList) {
        Clock loadClock = Clock.systemUTC();
        long startMills = loadClock.millis();
        log.info("flux-excel loading export template");
        try {
            //读所有的文件
            ExcelTemplateLoader.loadBootExportTemplate(fileList, EXCEL_TEMPLATE_CACHE_MAP, EXCEL_TEMPLATE_TITLE_CACHE_MAP);
        }catch (Exception e){
            log.error(Throwables.getStackTraceAsString(e));
            throw new ExcelReaderException("template.load.fail");
        }
        long loadMillis = loadClock.millis() - startMills;
        log.info("flux-excel load template success by {} millisecond", loadMillis);
    }

    /**
     *  导入模板解析并缓存
     */
    public static void loadImportExcelTemplate(List<Map<String, Object>> fileList, AbstractReaderTemplateExclude excluder) {
        Clock loadClock = Clock.systemUTC();
        long startMills = loadClock.millis();
        log.info("flux-excel loading import template");
        try {
            //读所有的文件
            ExcelTemplateLoader.loadBootImportTemplate(fileList,EXCEL_IMPORT_TEMPLATE_FILE_CACHE_MAP,EXCEL_IMPORT_TEMPLATE_CACHE_MAP,excluder);
        }catch (Exception e){
            log.error(Throwables.getStackTraceAsString(e));
            throw new ExcelReaderException("template.load.fail");
        }
        long loadMillis = loadClock.millis() - startMills;
        log.info("flux-excel load template success by {} millisecond", loadMillis);
    }

    /**
     * 获取模板文件list
     * @param files
     * @return
     */
    private static List<File> getFileList(File[] files) {
        List<File> fileList = new ArrayList<>();
        if (Objects.isNull(files) || files.length == ExcelConstant.ZERO_SHORT) {
            return fileList;
        }
        for (File file : files) {
            String templateName = file.getName();
            if (templateName.endsWith(ExcelConstant.XLSX_STR) || templateName.endsWith(ExcelConstant.XLS_STR)) {
                fileList.add(file);
            }
        }
        return fileList;
    }

    private static void loadExportTemplateIfNecessary(String templateName) {
        if (StringUtil.isEmpty(templateName) || EXCEL_TEMPLATE_CACHE_MAP.containsKey(templateName)) {
            return;
        }
        synchronized (EXPORT_TEMPLATE_LOAD_LOCK) {
            if (EXCEL_TEMPLATE_CACHE_MAP.containsKey(templateName)) {
                return;
            }
            List<Map<String, Object>> resources = resolveTemplateResources(EXPORT_TEMPLATE_DIR_PATH, templateName);
            if (!resources.isEmpty()) {
                loadExcelTemplate(resources);
            }
        }
    }

    private static void loadImportTemplateIfNecessary(String templateName) {
        if (StringUtil.isEmpty(templateName) || EXCEL_IMPORT_TEMPLATE_CACHE_MAP.containsKey(templateName)) {
            return;
        }
        synchronized (IMPORT_TEMPLATE_LOAD_LOCK) {
            if (EXCEL_IMPORT_TEMPLATE_CACHE_MAP.containsKey(templateName)) {
                return;
            }
            List<Map<String, Object>> resources = resolveTemplateResources(IMPORT_TEMPLATE_DIR_PATH, templateName);
            if (!resources.isEmpty()) {
                loadImportExcelTemplate(resources, IMPORT_TEMPLATE_EXCLUDER);
            }
        }
    }

    private static List<Map<String, Object>> resolveTemplateResources(String templateDir, String templateName) {
        List<Map<String, Object>> resources = new ArrayList<>();
        if (StringUtil.isEmpty(templateDir) || StringUtil.isEmpty(templateName)) {
            return resources;
        }
        if (Objects.nonNull(TEMPLATE_RESOURCE_RESOLVER)) {
            resources.addAll(Optional.ofNullable(TEMPLATE_RESOURCE_RESOLVER.resolve(templateDir, templateName)).orElse(Collections.emptyList()));
            if (!resources.isEmpty()) {
                return resources;
            }
        }
        File templateFile = new File(templateDir, templateName);
        if (!templateFile.exists() && templateDir.endsWith(ExcelConstant.FILE_SEPARATOR)) {
            templateFile = new File(templateDir + templateName);
        }
        if (!templateFile.isFile()) {
            return resources;
        }
        try {
            Map<String, Object> map = Maps.newHashMap();
            map.put("name", templateFile.getName());
            map.put("input", new FileInputStream(templateFile));
            map.put("cacheInput", new FileInputStream(templateFile));
            resources.add(map);
        } catch (FileNotFoundException e) {
            log.error("Load template file failed,cause:{}", Throwables.getStackTraceAsString(e));
            throw new ExcelReaderException("template.load.fail");
        }
        return resources;
    }

    private static void validateModelClass(Class key, java.util.function.Function<String, RuntimeException> exceptionFactory) {
        if (!ExcelBaseModel.class.isAssignableFrom(key)) {
            throw exceptionFactory.apply("Model class must extend ExcelBaseModel: " + key.getName());
        }
    }

}
