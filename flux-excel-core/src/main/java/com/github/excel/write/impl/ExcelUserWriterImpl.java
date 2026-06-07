package com.github.excel.write.impl;

import com.github.excel.boot.ExcelBootLoader;
import com.github.excel.constant.ExcelErrorMsgConstant;
import com.github.excel.enums.ExcelSuffixEnum;
import com.github.excel.exception.ExcelWriterException;
import com.github.excel.helper.ExcelValidationHelper;
import com.github.excel.helper.ExcelWriterHelper;
import com.github.excel.model.ExcelBaseModel;
import com.github.excel.model.ExcelWriterModel;
import com.github.excel.param.ExcelWriterCellParam;
import com.github.excel.param.ExcelWriterCommentParam;
import com.github.excel.param.ExcelWriterComboParam;
import com.github.excel.param.ExcelWriterConditionalStyleParam;
import com.github.excel.param.ExcelWriterDataParam;
import com.github.excel.param.ExcelWriterFileParam;
import com.github.excel.param.ExcelWriterListParam;
import com.github.excel.param.ExcelWriterMergeParam;
import com.github.excel.param.ExcelWriterModelParam;
import com.github.excel.param.ExcelWriterNumberScopeParam;
import com.github.excel.param.ExcelWriterParam;
import com.github.excel.param.ExcelWriterSteamParam;
import com.github.excel.util.ExcelUtil;
import com.github.excel.util.StringUtil;
import com.github.excel.write.BaseExcelWriter;
import com.github.excel.write.ExcelCustomWriter;
import com.github.excel.write.ExcelWriteKernel;
import com.github.excel.write.ExcelWriter;
import com.github.excel.write.pipeline.ExcelWriteContext;
import com.github.excel.write.pipeline.ExcelWritePipelines;
import com.github.excel.write.style.AbstractExcelStyle;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;

import jakarta.servlet.http.HttpServletRequest;
import jakarta.servlet.http.HttpServletResponse;
import jakarta.validation.groups.Default;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.Arrays;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.HashSet;
import java.util.LinkedHashMap;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.function.Function;
import java.util.stream.Collectors;
import java.util.stream.Stream;

/**
 * User-mode writer implementation that collects fluent API state and executes
 * the explicit write pipeline.
 */
@Slf4j
public class ExcelUserWriterImpl extends BaseExcelWriter implements ExcelWriter {

    private final ExcelWriterParam writerParam;
    private Class<? extends ExcelBaseModel> currentModelClass;
    private final Map<String, String> runtimeHeaderAliases = new LinkedHashMap<>();
    private final Set<String> runtimeIncludeFields = new LinkedHashSet<>();
    private final Map<Class<?>, Function<Object, Object>> runtimeTypeConverters = new LinkedHashMap<>();
    private final Map<String, Function<Object, Object>> runtimeFieldConverters = new LinkedHashMap<>();
    private boolean runtimeOnlyAlias;

    public ExcelUserWriterImpl(ExcelWriterParam writerParam) {
        this.writerParam = writerParam;
        this.streaming = Boolean.TRUE.equals(writerParam.getStreaming());
        this.template = writerParam.getTemplate();
        this.selectSheet = writerParam.getSelectSheetName();
        this.noneDataTips = Boolean.TRUE.equals(writerParam.getNoneDataTips());
        if (writerParam.getSheetRowMaxCount() != null) {
            this.sheetRowMaxCount = writerParam.getSheetRowMaxCount();
        }
    }

    @Override
    public <T extends ExcelBaseModel> ExcelWriter writeModel(ExcelWriterModelParam<T> param) {
        validateParam(param);
        ExcelWriterHelper.excludes(param.getModelCla(), param.getExcludeFields(), excludeFieldMap);
        currentModelClass = param.getModelCla();
        exportBeanList.add(createWriterModel(param.getModelCla(), param.getModel(), null, param));
        return this;
    }

    @Override
    public <T extends ExcelBaseModel> ExcelWriter writeList(ExcelWriterListParam<T> param) {
        validateParam(param);
        ExcelWriterHelper.excludes(param.getModelCla(), param.getExcludeFields(), excludeFieldMap);
        listCla = param.getModelCla();
        currentModelClass = param.getModelCla();
        exportModelList.add(createWriterModel(param.getModelCla(), null, param.getModelList(), param));
        return this;
    }

    @Override
    public ExcelWriter writeColumn(ExcelWriterCellParam customColumnModel) {
        validateParam(customColumnModel);
        customColumnModelList.add(customColumnModel);
        return this;
    }

    @Override
    public ExcelWriter writeMergeColumn(ExcelWriterMergeParam mergeCustomColumnModel) {
        validateParam(mergeCustomColumnModel);
        mergeCustomColumnModelList.add(mergeCustomColumnModel);
        return this;
    }

    @Override
    public ExcelWriter writeCustom(ExcelCustomWriter customWrite) {
        this.customWrite = customWrite;
        return this;
    }

    @Override
    public <T extends AbstractExcelStyle> ExcelWriter addStyles(Class<T>... styles) {
        HashSet<Class<T>> collect = Stream.of(styles).collect(Collectors.toCollection(HashSet::new));
        styleList.addAll(collect);
        return this;
    }

    @Override
    public CellStyle getStyle(String name) {
        return styleLocal.get().get(name);
    }

    @Override
    public Font getFont(String name) {
        return fontLocal.get().get(name);
    }

    @Override
    public ExcelWriter export() {
        if (writerParam instanceof ExcelWriterSteamParam) {
            ExcelWriterSteamParam steamParam = (ExcelWriterSteamParam) writerParam;
            String fileName = "export" + writerParam.getSuffixEnum().getSuffix();
            export(steamParam.getOutputStream(), fileName, writerParam.getSuffixEnum());
            return this;
        }
        if (writerParam instanceof ExcelWriterFileParam) {
            ExcelWriterFileParam fileParam = (ExcelWriterFileParam) writerParam;
            String fileName = fileParam.getFile().getName();
            try (OutputStream outputStream = new FileOutputStream(fileParam.getFile())) {
                export(outputStream, fileName, writerParam.getSuffixEnum());
            } catch (IOException e) {
                throw new ExcelWriterException(e);
            }
            return this;
        }
        throw new ExcelWriterException("Unsupported writer param");
    }

    @Override
    public ExcelWriter export(OutputStream outputStream, String fileName, ExcelSuffixEnum suffixEnum) {
        String excelName = fileName.endsWith(suffixEnum.getSuffix()) ? fileName : fileName + suffixEnum.getSuffix();
        writerParam.setSuffixEnum(suffixEnum);
        ExcelWriteContext context = ExcelWriteContext.builder()
                .outputStream(outputStream)
                .excelName(excelName)
                .writerParam(writerParam)
                .streaming(streaming)
                .noneDataTips(noneDataTips)
                .selectSheet(selectSheet)
                .listCla(listCla)
                .sheetRowMaxCount(sheetRowMaxCount)
                .customWrite(customWrite)
                .exportModelList(exportModelList)
                .exportBeanList(exportBeanList)
                .customColumnModelList(customColumnModelList)
                .mergeCustomColumnModelList(mergeCustomColumnModelList)
                .styleList(styleList)
                .excludeFieldMap(excludeFieldMap)
                .commentMap(commentMap)
                .conditionalStyleMap(conditionalStyleMap)
                .conditionalStyleList(conditionalStyleList)
                .runtimeHeaderAliases(runtimeHeaderAliases)
                .runtimeIncludeFields(runtimeIncludeFields)
                .runtimeOnlyAlias(runtimeOnlyAlias)
                .runtimeTypeConverters(runtimeTypeConverters)
                .runtimeFieldConverters(runtimeFieldConverters)
                .build();
        ExcelWritePipelines.workbookPipeline(new ExcelWriteKernel()).execute(context);
        clearData();
        return this;
    }

    @Override
    public ExcelWriter export(HttpServletRequest request, HttpServletResponse response, String fileName, ExcelSuffixEnum suffixEnum) {
        try {
            ExcelUtil.setResponseHeader(request, response, fileName, suffixEnum.getSuffix());
            export(response.getOutputStream(), fileName, suffixEnum);
            return this;
        } catch (IOException e) {
            throw new ExcelWriterException(e);
        }
    }

    @Override
    public ExcelWriter process(OutputStream outputStream, String fileName, ExcelSuffixEnum suffixEnum) {
        return export(outputStream, fileName, suffixEnum);
    }

    @Override
    public ExcelWriter process(HttpServletRequest request, HttpServletResponse response, String fileName, ExcelSuffixEnum suffixEnum) {
        return export(request, response, fileName, suffixEnum);
    }

    @Override
    public ExcelWriter process(File file, String fileName, ExcelSuffixEnum suffixEnum) {
        try (OutputStream outputStream = new FileOutputStream(file)) {
            return export(outputStream, fileName, suffixEnum);
        } catch (IOException e) {
            throw new ExcelWriterException(e);
        }
    }

    @Override
    public ExcelWriter excludes(String[] excludeFields) {
        if (currentModelClass == null || excludeFields == null) {
            return this;
        }
        ExcelWriterHelper.excludes(currentModelClass, Arrays.asList(excludeFields), excludeFieldMap);
        return this;
    }

    @Override
    public ExcelWriter addValidationOrComment(String field, ExcelWriterCommentParam commentParam) {
        if (currentModelClass == null || StringUtil.isEmpty(field) || commentParam == null) {
            return this;
        }
        Map<String, ExcelWriterCommentParam> fieldMap = commentMap.computeIfAbsent(currentModelClass, key -> new HashMap<>());
        fieldMap.put(field, commentParam);
        return this;
    }

    @Override
    public ExcelWriter addValidationOrComment(String field, ExcelWriterNumberScopeParam validationParam) {
        if (validationParam == null) {
            return this;
        }
        ExcelWriterCommentParam commentParam = new ExcelWriterCommentParam();
        commentParam.setValue(validationParam.getValue());
        commentParam.setCommentText(validationParam.getMessage());
        commentParam.setAuthor(validationParam.getTitle());
        return addValidationOrComment(field, commentParam);
    }

    @Override
    public ExcelWriter addValidationOrComment(String field, ExcelWriterComboParam validationParam) {
        if (validationParam == null) {
            return this;
        }
        ExcelWriterCommentParam commentParam = new ExcelWriterCommentParam();
        commentParam.setValue(validationParam.getValue());
        commentParam.setCommentText(validationParam.getMessage());
        commentParam.setAuthor(validationParam.getTitle());
        return addValidationOrComment(field, commentParam);
    }

    @Override
    public ExcelWriter addConditionalStyle(String field, ExcelWriterConditionalStyleParam conditionalStyleParam) {
        if (currentModelClass == null || StringUtil.isEmpty(field) || conditionalStyleParam == null) {
            return this;
        }
        Map<String, List<ExcelWriterConditionalStyleParam>> fieldMap = conditionalStyleMap.computeIfAbsent(currentModelClass, key -> new HashMap<>());
        fieldMap.computeIfAbsent(field, key -> new ArrayList<>()).add(conditionalStyleParam);
        return this;
    }

    @Override
    public ExcelWriter addConditionalStyle(ExcelWriterConditionalStyleParam conditionalStyleParam) {
        if (conditionalStyleParam == null) {
            return this;
        }
        conditionalStyleList.add(conditionalStyleParam);
        return this;
    }

    @Override
    public ExcelWriter setNoneDataTips(boolean noneDataTips) {
        this.noneDataTips = noneDataTips;
        writerParam.setNoneDataTips(noneDataTips);
        return this;
    }

    @Override
    public ExcelWriter setStreaming(boolean streaming) {
        this.streaming = streaming;
        writerParam.setStreaming(streaming);
        return this;
    }

    @Override
    public ExcelWriter selectSheet(String sheetName) {
        this.selectSheet = sheetName;
        writerParam.setSelectSheetName(sheetName);
        return this;
    }

    @Override
    public ExcelWriter setListCla(Class<? extends ExcelBaseModel> listCla) {
        this.listCla = listCla;
        this.currentModelClass = listCla;
        return this;
    }

    @Override
    public ExcelWriter runtimeHeaderAliases(Map<String, String> aliases) {
        runtimeHeaderAliases.clear();
        if (aliases != null) {
            runtimeHeaderAliases.putAll(aliases);
        }
        return this;
    }

    @Override
    public ExcelWriter runtimeIncludeFields(Set<String> fields) {
        runtimeIncludeFields.clear();
        if (fields != null) {
            runtimeIncludeFields.addAll(fields);
        }
        return this;
    }

    @Override
    public ExcelWriter runtimeOnlyAlias(boolean onlyAlias) {
        this.runtimeOnlyAlias = onlyAlias;
        return this;
    }

    @Override
    public ExcelWriter runtimeTypeConverters(Map<Class<?>, Function<Object, Object>> converters) {
        runtimeTypeConverters.clear();
        if (converters != null) {
            runtimeTypeConverters.putAll(converters);
        }
        return this;
    }

    @Override
    public ExcelWriter runtimeFieldConverters(Map<String, Function<Object, Object>> converters) {
        runtimeFieldConverters.clear();
        if (converters != null) {
            runtimeFieldConverters.putAll(converters);
        }
        return this;
    }

    private ExcelWriterModel createWriterModel(Class<? extends ExcelBaseModel> modelClass,
                                               ExcelBaseModel model,
                                               java.util.List<? extends ExcelBaseModel> modelList,
                                               ExcelWriterDataParam param) {
        ExcelWriterModel writerModel = new ExcelWriterModel();
        writerModel.setSheetName(param.getSheetName());
        writerModel.setDataModel(model);
        writerModel.setDataModelList(modelList);
        writerModel.setCacheModel(ExcelBootLoader.getExcelCacheMapValue(modelClass));
        writerModel.setExcelModelClass(modelClass);
        writerModel.setRowIndex(param.getRowIndex());
        writerModel.setColIndex(param.getColIndex());
        writerModel.setFillTemplate(param.getFillTemplate());
        return writerModel;
    }

    private void validateParam(ExcelWriterDataParam param) {
        ExcelValidationHelper.validationBean(param, new Class[]{Default.class});
        if (StringUtil.isEmpty(param.getSheetName())) {
            throw new ExcelWriterException(ExcelErrorMsgConstant.ERROR_SHEET_NAME);
        }
    }

    @Override
    protected void clearData() {
        super.clearData();
        currentModelClass = null;
        runtimeHeaderAliases.clear();
        runtimeIncludeFields.clear();
        runtimeTypeConverters.clear();
        runtimeFieldConverters.clear();
        runtimeOnlyAlias = false;
    }
}
