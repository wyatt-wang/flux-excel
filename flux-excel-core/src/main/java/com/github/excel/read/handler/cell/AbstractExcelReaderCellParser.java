package com.github.excel.read.handler.cell;

import com.github.excel.context.ExcelReaderContext;
import com.github.excel.context.ExcelReaderModelContext;
import com.github.excel.model.ExcelBaseModel;
import com.github.excel.model.ExcelCacheImportModel;
import com.github.excel.read.config.ReadModelListTitleFieldConfig;
import com.github.excel.read.config.ReadModelTitleConfig;
import com.google.common.collect.Maps;

import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.stream.Collectors;

/**
 * @Author: Vachel Wang
 * @Date: 2026/4/24
 * @Description: 批量处理大小
 */
public abstract class AbstractExcelReaderCellParser<T extends ExcelBaseModel> implements ExcelReaderCellParser<T> {

    protected Map<Class<T>, ReadModelTitleConfig<T>> titleConfigMap = Maps.newConcurrentMap();

    public abstract void beforeParse(ExcelReaderContext<T> readerContext);

    public abstract void afterParse(ExcelReaderContext<T> readerContext);

    public abstract void parseBean(ExcelReaderContext<T> readerContext);

    public abstract void parseBeanList(ExcelReaderContext<T> readerContext);

    @Override
    public void cellParser(ExcelReaderContext<T> readerContext) {
        // 填充之前
        beforeParse(readerContext);
        // 过滤需要填充的bean
        List<ExcelReaderModelContext<T>> fillModels = readerContext.getModelMap().values().stream().filter(e -> e.getCacheImportModel().getFieldModelMap().get(readerContext.getParserContext().getCellValue()) != null
                        && e.getModelList() == null)
                .collect(Collectors.toList());
        // 过滤需要填充的bean list
        List<ExcelReaderModelContext<T>> fillModelLists = readerContext.getModelMap().values().stream().filter(e -> e.getModelList() != null)
                .collect(Collectors.toList());
        readerContext.getParserContext().setFillModels(fillModels);
        readerContext.getParserContext().setFillModelLists(fillModelLists);

        parseBean(readerContext);
        parseBeanList(readerContext);
        afterParse(readerContext);
    }

    /**
     * 构建标题映射map
     * @return
     */
    protected ReadModelTitleConfig<T> buildTitleConfigMap(Class<T> clazz, Integer startRow, Integer colIndex, ExcelCacheImportModel.ExcelCacheImportFieldModel fieldModel, String titleName, Integer rowIndex) {
        ReadModelTitleConfig<T> readModelTitleConfig = titleConfigMap.get(clazz);
        ReadModelListTitleFieldConfig fieldConfig = new ReadModelListTitleFieldConfig(fieldModel, titleName, rowIndex);
        if (Objects.isNull(readModelTitleConfig)) {
            readModelTitleConfig = new ReadModelTitleConfig<>();
            readModelTitleConfig.setClazz(clazz);
            readModelTitleConfig.getFieldConfigMap().put(colIndex, fieldConfig);
            readModelTitleConfig.setStartRow(startRow);
            titleConfigMap.put(clazz, readModelTitleConfig);
        }else{
            readModelTitleConfig.getFieldConfigMap().put(colIndex, fieldConfig);
        }
        return readModelTitleConfig;
    }
}
