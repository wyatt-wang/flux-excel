package com.github.excel.read.handler.cell;

import com.github.excel.constant.ExcelConstant;
import com.github.excel.constant.ExcelErrorMsgConstant;
import com.github.excel.context.ExcelReaderContext;
import com.github.excel.context.ExcelReaderModelContext;
import com.github.excel.exception.ExcelReaderException;
import com.github.excel.helper.ExcelConvertDataHelper;
import com.github.excel.helper.ExcelHelper;
import com.github.excel.helper.ExcelValidationHelper;
import com.github.excel.model.ExcelBaseModel;
import com.github.excel.model.ExcelCacheImportModel;
import com.github.excel.model.ExcelReadError;
import com.github.excel.model.ExcelReaderPictureModel;
import com.github.excel.param.ExcelReaderListParam;
import com.github.excel.read.config.ReadModelTitleConfig;
import com.github.excel.read.config.ReadModelListTitleFieldConfig;
import com.github.excel.read.format.ExcelReaderFormatManager;
import com.github.excel.util.StringUtil;
import com.google.common.base.Throwables;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.collections4.CollectionUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellReference;

import java.lang.reflect.InvocationTargetException;
import java.text.ParseException;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.function.Function;
import java.util.stream.Collectors;

/**
 * Excel 单元格标准解析
 * 需要匹配表头，并且匹配完成之后不再匹配，匹配后需记录坐标，当列相同，并且row大于的时候进行设值，表头要求在同一行
 * 如果当前行映射不了则跳过（需要记录class），继续下一行
 * 1、按照class来存储，class里面field的读取坐标值（开始行，读取列）
 * 2、里面的映射完成后就不再执行映射
 * 3、未能匹配到表头判断
 * 4、匹配到任意一个就算是title行
 * @author Vico
 * @create 2022-12-14 17:38
 */
@Slf4j
public class ExcelReaderCellParserImpl<T extends ExcelBaseModel> extends AbstractExcelReaderCellParser<T> {
    @Override
    public void beforeParse(ExcelReaderContext<T> readerContext) {

    }

    @Override
    public void afterParse(ExcelReaderContext<T> readerContext) {

    }

    @Override
    public void parseBean(ExcelReaderContext<T> readerContext) {
        List<ExcelReaderModelContext<T>> readModelList = readerContext.getParserContext().getFillModels();
        Object cellValue = readerContext.getParserContext().getCellValue();
        Cell cell = readerContext.getParserContext().getCell();
        if (CollectionUtils.isEmpty(readModelList) || Objects.isNull(cellValue)) {
            return;
        }
        for (ExcelReaderModelContext<T> modelDto : readModelList) {
            ExcelCacheImportModel.ExcelCacheImportFieldModel fieldModel = modelDto.getCacheImportModel().getFieldModelMap().get(cellValue);
            if (Objects.isNull(fieldModel)) {
                continue;
            }
            try {
                // 获取title配置
                ReadModelTitleConfig<T> listTitleConfig = buildTitleConfigMap(modelDto.getModelCla(), cell.getRowIndex(), (cell.getColumnIndex() + ExcelConstant.ONE_INT), fieldModel, (String) cellValue, cell.getRowIndex());
                modelDto.setListTitleConfig(listTitleConfig);
                Class<?> parameterType = fieldModel.getSetMethod().getParameterTypes()[ExcelConstant.ZERO_SHORT];
                CellReference cellReference = new CellReference(cell.getRowIndex(), cell.getColumnIndex() + ExcelConstant.ONE_INT);
                String cellReferenceStr = cellReference.formatAsString();
                Object setParams = ExcelHelper.getCellValue(cell.getRow().getCell(cell.getColumnIndex() + ExcelConstant.ONE_INT), parameterType, cellReferenceStr, readerContext.getParserContext().getReaderFormatManager().getFormulaEvaluator());
                // get picture value
                Object setPicParams = ExcelHelper.getPictureValue(readerContext.getParserContext().getSheetPictureMap(), cell.getRowIndex(), cell.getColumnIndex() + ExcelConstant.ONE_INT, setParams, parameterType);

                modelDto.getModel().getModelColAddress().put(fieldModel.getField(), cellReferenceStr);
                // 判断并进行转换
                if (setPicParams == setParams) {
                    setParams = ExcelConvertDataHelper.convertDataType(setParams, fieldModel);
                }else{
                    setParams = setPicParams;
                }
                setParams = applyUserConverter(readerContext, fieldModel, setParams);
                // 校验为空数据
                if (fieldModel.getImportProperty().checkNull() &&
                        StringUtil.isEmpty(setParams)) {
                    String errorTips = String.format(ExcelErrorMsgConstant.ERROR_DATA_NULL_READ_MSG, cell.getSheet().getSheetName(), cellReference.formatAsString(), fieldModel.getImportProperty().titleName());
                    throw new ExcelReaderException(errorTips);
                }

                if (Objects.isNull(setParams)) {
                    continue;
                }

                // 格式化数据
                setParams = readerContext.getParserContext().getReaderFormatManager().getDataFormatThenCache(fieldModel.getImportProperty().formatter()).format(setParams, fieldModel.getImportProperty().formatPattern(), parameterType);
                // 填充数据
                fieldModel.getSetMethod().invoke(modelDto.getModel(), setParams);
            } catch (Exception e) {
                Cell valueCell = cell.getRow().getCell(cell.getColumnIndex() + ExcelConstant.ONE_INT);
                if (!recordError(readerContext, Objects.nonNull(valueCell) ? valueCell : cell, fieldModel, e)) {
                    if (e instanceof ExcelReaderException) {
                        throw (ExcelReaderException) e;
                    }
                    log.error("read excel failure,cause:{}",Throwables.getStackTraceAsString(e));
                    throw new ExcelReaderException(e.getMessage());
                }
            }

        }
    }

    @Override
    public void parseBeanList(ExcelReaderContext<T> readerContext) {
        List<ExcelReaderModelContext<T>> readModelList = readerContext.getParserContext().getFillModelLists();
        if (CollectionUtils.isEmpty(readModelList)) {
            return;
        }
        // configTitle
        modelListTitleConfig(readerContext);
        // 解析内容
        readModelList.forEach(readModel->{
            // 循环解析
            ReadModelTitleConfig<T> listTitleConfig = titleConfigMap.get(readModel.getModelCla());
            if(listTitleConfig != null) {
                if(Objects.isNull(listTitleConfig.getLastCol())) {
                    Integer lastCol = listTitleConfig.getFieldConfigMap().keySet().stream().max(Integer::compareTo).orElse(ExcelConstant.ONE_INT);
                    listTitleConfig.setLastCol(lastCol);
                }
                // 进行设值
                readerContext.getParserContext().setListTitleConfig(listTitleConfig);
                readerContext.getParserContext().setCurrentModelContext(readModel);
                parseBeanListField(readerContext);
            }
        });

    }


    /**
     * 解析并设置Bean字段
     */
    private void parseBeanListField(ExcelReaderContext<T> readerContext) {
        ReadModelTitleConfig<T> listTitleConfig = readerContext.getParserContext().getListTitleConfig();
        Cell cell = readerContext.getParserContext().getCell();
        ExcelReaderModelContext<T> readModel = readerContext.getParserContext().getCurrentModelContext();
        ExcelReaderFormatManager readerFormatManager = readerContext.getParserContext().getReaderFormatManager();
        Map<String, List<ExcelReaderPictureModel>> sheetPictureMap = readerContext.getParserContext().getSheetPictureMap();
        if(cell.getRowIndex() <= listTitleConfig.getStartRow()){
            return ;
        }
        // 行号跳跃直接return
        if (Objects.nonNull(readModel.getRowNum()) && readModel.getRowNum() < cell.getRowIndex() - ExcelConstant.ONE_INT) {
            return ;
        }
        ReadModelListTitleFieldConfig fieldConfig = listTitleConfig.getFieldConfigMap().get(cell.getColumnIndex());
        if (Objects.isNull(fieldConfig)) {
            return ;
        }
        readModel.setRowNum(cell.getRowIndex());
        CellReference cellReference = new CellReference(cell.getRowIndex(), cell.getColumnIndex());
        ExcelCacheImportModel.ExcelCacheImportFieldModel importFieldModel = readModel.getCacheImportModel().getFieldModelMap().get(fieldConfig.getTitleName());
        Class<?> parameterType = fieldConfig.getCacheImportFieldModel().getSetMethod().getParameterTypes()[ExcelConstant.ZERO_SHORT];
        // 获取值
        Object value = ExcelHelper.getCellValue(cell, parameterType,cellReference.formatAsString(),readerFormatManager.getFormulaEvaluator());
        if (StringUtil.isEmpty(value) && Objects.nonNull(readerContext.getParserContext().getCellValue())) {
            value = readerContext.getParserContext().getCellValue();
        }
        if (StringUtil.isEmpty(value) && readerContext.getParserContext().getMergedCellValueMap() != null) {
            value = readerContext.getParserContext().getMergedCellValueMap()
                    .get(cell.getRowIndex() + ":" + cell.getColumnIndex());
        }

        T baseModel = getOrCreateModel(cell.getRowIndex(), listTitleConfig, readModel);
        baseModel.getModelColAddress().put(importFieldModel.getField(), cellReference.formatAsString());
        // 到达最后一列执行校验，坐标重新设置
        if(cell.getColumnIndex() == cell.getRow().getLastCellNum() - ExcelConstant.ONE_INT && cell.getCellType() == CellType.BLANK) {
            executeCallback(cell, readModel, listTitleConfig, baseModel);
        }else if(cell.getColumnIndex() == listTitleConfig.getLastCol() && cell.getCellType() == CellType.BLANK){
            executeCallback(cell, readModel, listTitleConfig, baseModel);
        }
        // 获取图片值
        Object picValue = ExcelHelper.getPictureValue(sheetPictureMap, cell.getRowIndex(), cell.getColumnIndex(), value, parameterType);
        // 转换值
        if(Objects.nonNull(value) && value == picValue) {
            value = ExcelConvertDataHelper.convertDataType(value, importFieldModel);
        }else{
            value = picValue;
        }
        // 校验为空数据
        if (importFieldModel.getImportProperty().checkNull() &&
                StringUtil.isEmpty(value)) {
            String errorTips = String.format(ExcelErrorMsgConstant.ERROR_DATA_NULL_READ_MSG, cell.getSheet().getSheetName(), cellReference.formatAsString(), importFieldModel.getImportProperty().titleName());
            throw new ExcelReaderException(errorTips);
        }

        if (StringUtil.isEmpty(value)) {
            return ;
        }

        try {
            value = applyUserConverter(readerContext, importFieldModel, value);
            value = readerFormatManager.getDataFormatThenCache(importFieldModel.getImportProperty().formatter()).format(value, importFieldModel.getImportProperty().formatPattern(), parameterType);
        } catch (Exception e) {
            if (recordError(readerContext, cell, importFieldModel, e)) {
                return;
            }
            log.error(Throwables.getStackTraceAsString(e));
            String errorMsg = String.format("解析Excel失败，原因：格式化失败，单元格：%s，数据:%s", cellReference.formatAsString(), value);
            throw new ExcelReaderException(errorMsg);
        }
        try {
            fieldConfig.getCacheImportFieldModel().getSetMethod().invoke(baseModel, value);
            // 到达最后一列执行校验，坐标重新设置
            if(cell.getColumnIndex() == cell.getRow().getLastCellNum() - ExcelConstant.ONE_INT) {
                executeCallback(cell, readModel, listTitleConfig, baseModel);
            }else if(cell.getColumnIndex() == listTitleConfig.getLastCol()) {
                executeCallback(cell, readModel, listTitleConfig, baseModel);
            }
        } catch (IllegalAccessException | InvocationTargetException e) {
            if (recordError(readerContext, cell, importFieldModel, e)) {
                return;
            }
            log.error(Throwables.getStackTraceAsString(e));
            String errorMsg = String.format("解析Excel失败，原因：解析数据失败，单元格：%s，数据:%s", cellReference.formatAsString(), value);
            throw new ExcelReaderException(errorMsg);
        }
    }

    private Object applyUserConverter(ExcelReaderContext<T> readerContext,
                                      ExcelCacheImportModel.ExcelCacheImportFieldModel fieldModel,
                                      Object value) {
        Function<Object, Object> fieldConverter = readerContext.getReaderParam().getFieldConverters().get(fieldModel.getField());
        if (fieldConverter != null) {
            return fieldConverter.apply(value);
        }
        if (value == null) {
            return null;
        }
        for (Map.Entry<Class<?>, Function<Object, Object>> entry : readerContext.getReaderParam().getTypeConverters().entrySet()) {
            if (entry.getKey().isAssignableFrom(value.getClass())) {
                return entry.getValue().apply(value);
            }
        }
        return value;
    }

    private boolean recordError(ExcelReaderContext<T> readerContext,
                                Cell cell,
                                ExcelCacheImportModel.ExcelCacheImportFieldModel fieldModel,
                                Exception exception) {
        if (!Boolean.TRUE.equals(readerContext.getReaderParam().getCollectErrors())) {
            return false;
        }
        ExcelReadError error = new ExcelReadError()
                .setSheetIndex(cell.getSheet().getWorkbook().getSheetIndex(cell.getSheet()))
                .setSheetName(cell.getSheet().getSheetName())
                .setRowIndex(cell.getRowIndex())
                .setColIndex(cell.getColumnIndex())
                .setFieldName(fieldModel.getField())
                .setTitleName(fieldModel.getImportProperty().titleName())
                .setRawValue(readRawCellValue(cell))
                .setMessage(exception.getMessage());
        readerContext.getReaderParam().getReadErrors().add(error);
        return true;
    }

    private Object readRawCellValue(Cell cell) {
        try {
            return ExcelHelper.getCellValue(cell, String.class, cell.getAddress().formatAsString(), null);
        } catch (Exception ignored) {
            return null;
        }
    }

    private void executeCallback(Cell cell, ExcelReaderModelContext<T> readModel,
                                 ReadModelTitleConfig<T> listTitleConfig, T baseModel
                                 ) {
        listTitleConfig.getFieldConfigMap().forEach((k, v)->{v.setRowIndex(cell.getRowIndex());});
        ExcelValidationHelper.resetBeanColAddress(baseModel, listTitleConfig);
        ExcelValidationHelper.validationBean(readModel, baseModel);
        if(Objects.nonNull(readModel.getParam().getRowHandler())){
            readModel.getParam().getRowHandler().  handler(baseModel);
        }
        // 执行回调
        ExcelReaderListParam<T> listParam =(ExcelReaderListParam<T>) readModel.getParam();
        if (listParam.getBatchProcess() != null) {
            listParam.getBatchProcess().doProcess(readModel.getModelList());
        }
    }


    private T getOrCreateModel(int rowIndex, ReadModelTitleConfig<T> listTitleConfig, ExcelReaderModelContext<T> readModel) {
        T baseModel = listTitleConfig.getRowBeanMap().get(rowIndex);
        if (baseModel == null) {
            try {
                baseModel = listTitleConfig.getClazz().newInstance();
                readModel.getModelList().add(baseModel);
            } catch (InstantiationException | IllegalAccessException e) {
                log.error(Throwables.getStackTraceAsString(e));
                throw new ExcelReaderException(e.getMessage());
            }
            listTitleConfig.getRowBeanMap().put(rowIndex, baseModel);
            return baseModel;
        } else {
            return baseModel;
        }
    }

    private void modelListTitleConfig(ExcelReaderContext<T> readerContext) {
        List<ExcelReaderModelContext<T>> readModelList = readerContext.getParserContext().getFillModelLists();
        Cell cell = readerContext.getParserContext().getCell();
        ExcelReaderFormatManager readerFormatManager = readerContext.getParserContext().getReaderFormatManager();
        // 如果已经存在则跳过
        long titleCount = readModelList.stream().filter(e -> Objects.isNull(titleConfigMap.get(e.getModelCla()))).count();
        if(titleCount < ExcelConstant.ONE_INT){
            return ;
        }
        Row titleRow = cell.getRow();
        int rowIndex = titleRow.getRowNum();
        // 从第一列开始解析
        int colIndex = titleRow.getFirstCellNum(), lastColIndex = titleRow.getLastCellNum();
        for (; colIndex < lastColIndex; colIndex++) {
            Cell titleCell = titleRow.getCell(colIndex);
            // 如果为空继续执行
            if (Objects.isNull(titleCell)) {
                continue;
            }
            Object title = ExcelHelper.getCellValue(titleCell, String.class, null,readerFormatManager.getFormulaEvaluator());
            if (StringUtil.isEmpty(title)) {
                continue;
            }
            // 需要导入的field
            List<ExcelReaderModelContext<T>> readModelDtoList = readModelList.stream().filter(e -> e.getCacheImportModel().getFieldModelMap().get(title) != null).collect(Collectors.toList());
            if (CollectionUtils.isEmpty(readModelDtoList)) {
                continue;
            }
            // 需要解析的dto list
            for (ExcelReaderModelContext<T> readModelDto : readModelDtoList) {
                ExcelCacheImportModel.ExcelCacheImportFieldModel fieldModel = readModelDto.getCacheImportModel().getFieldModelMap().get(title);
                ReadModelTitleConfig<T> listTitleConfig = buildTitleConfigMap(readModelDto.getModelCla(), rowIndex, colIndex, fieldModel, title.toString(), rowIndex);
                readModelDto.setListTitleConfig(listTitleConfig);
            }

        }
        // 映射完成后，有值的情况下，执行判断
        ExcelValidationHelper.checkTitle(readModelList);
    }


}
