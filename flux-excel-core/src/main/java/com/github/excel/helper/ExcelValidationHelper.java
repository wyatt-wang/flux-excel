package com.github.excel.helper;

import com.github.excel.context.ExcelReaderModelContext;
import com.github.excel.exception.ExcelReaderException;
import com.github.excel.model.ExcelBaseModel;
import com.github.excel.model.ExcelCacheImportModel;
import com.github.excel.validator.ExcelValidatorConfig;
import com.github.excel.read.config.ReadModelTitleConfig;
import com.github.excel.read.config.ReadModelListTitleFieldConfig;
import com.github.excel.util.StringUtil;
import org.apache.poi.ss.util.CellReference;

import jakarta.validation.ConstraintViolation;
import jakarta.validation.groups.Default;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.Set;
import java.util.concurrent.ConcurrentMap;
import java.util.stream.Collectors;

/**
 * 校验工具类
 * @author Vico
 * @create 2022-12-15 18:52
 */
public class ExcelValidationHelper {
    /**
     * 校验Bean
     * @param modelContext model
     * @param bean bean
     */
    public static <T extends ExcelBaseModel> void validationBean(ExcelReaderModelContext<T> modelContext , ExcelBaseModel bean) {
        bean.callback();
        if(Objects.nonNull(modelContext.getCacheImportModel().getValidation())) {
            Class<?>[] validateCla = bean.validationGroup();
            if(validateCla == null || validateCla.length == 0){
                validateCla = new Class[]{Default.class};
            }
            Set<ConstraintViolation<ExcelBaseModel>> validate = ExcelValidatorConfig.getValidator().validate(bean, validateCla);
            validate.forEach(e -> {
                String errorPoint = bean.getModelColAddress().get(e.getPropertyPath().toString()), errorMsg;
                if (Objects.isNull(errorPoint)) {
                    errorMsg = "解析Excel失败，字段为空：" + modelContext.getCacheImportModel().getFieldModelMap().entrySet().stream().filter(k -> k.getValue().getField().equals(e.getPropertyPath().toString())).map(j -> j.getValue().getImportProperty().titleName()).findFirst().orElseGet(() -> null) + "，并且" + e.getMessage();
                } else {
                    errorMsg = "解析Excel失败，坐标：" + bean.getModelColAddress().get(e.getPropertyPath().toString()) + "，原因：" + e.getMessage();
                }
                throw new ExcelReaderException(errorMsg);
            });
        }
    }

    public static <T extends ExcelBaseModel> void checkTitle(List<ExcelReaderModelContext<T>> readModelList){
        if (readModelList.isEmpty()) {
            return ;
        }
        for (ExcelReaderModelContext<T> readModelContext : readModelList) {
            if (Objects.isNull(readModelContext.getListTitleConfig()) || !readModelContext.getCacheImportModel().getExcelRead().checkTitle()) {
                continue;
            }
            Integer startRow = readModelContext.getListTitleConfig().getStartRow();
            if (Objects.isNull(startRow)) {
                continue;
            }
            // 匹配field
            ConcurrentMap<String, ExcelCacheImportModel.ExcelCacheImportFieldModel> titleConfigFieldMap = readModelContext.getListTitleConfig().getFieldConfigMap().entrySet().stream().map(Map.Entry::getValue).collect(Collectors.toConcurrentMap((key -> key.getCacheImportFieldModel().getField()), (ReadModelListTitleFieldConfig::getCacheImportFieldModel)));
            readModelContext.getCacheImportModel().getFieldModelMap().forEach((key, value) -> {
                ExcelCacheImportModel.ExcelCacheImportFieldModel importFieldModel = titleConfigFieldMap.get(value.getField());
                if (Objects.isNull(importFieldModel)) {
                    throw new ExcelReaderException("解析失败，未能找到表头：" + key);
                }
            });
        }
    }

    /**
     * 设置model字段对应的坐标信息
     * @param baseModel model
     * @param listTitleConfig 配置
     */
    public static <T extends ExcelBaseModel> void resetBeanColAddress(ExcelBaseModel baseModel , ReadModelTitleConfig<T> listTitleConfig){
        Map<String, String> modelColAddress = baseModel.getModelColAddress();
        if (modelColAddress.isEmpty()) {
            return ;
        }
        listTitleConfig.getFieldConfigMap().entrySet().stream().forEach(entry->{
            Integer colIndex = entry.getKey();
            ExcelCacheImportModel.ExcelCacheImportFieldModel field = entry.getValue().getCacheImportFieldModel();
            String address = modelColAddress.get(field.getField());
            if (StringUtil.isEmpty(address)) {
                CellReference cellReference = new CellReference(entry.getValue().getRowIndex(), colIndex);
                modelColAddress.put(field.getField(), cellReference.formatAsString());
            }
        });
    }

    public static void validationBean(Object bean, Class<?>[] validateCla) {
        if (Objects.isNull(bean)) {
            throw new ExcelReaderException("bean不能为空");
        }
        if(validateCla == null || validateCla.length == 0){
            validateCla = new Class[]{Default.class};
        }
        ExcelValidatorConfig.getValidator().validate(bean, validateCla).forEach(e->{
            throw new ExcelReaderException(e.getMessage());
        });
    }
}
