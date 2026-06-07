package com.github.excel.context;

import com.github.excel.model.ExcelBaseModel;
import com.github.excel.model.ExcelCacheImportModel;
import com.github.excel.param.ExcelReaderDataParam;
import com.github.excel.read.config.ReadModelTitleConfig;
import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;
import lombok.experimental.Accessors;

import java.util.List;

/**
 * Excel 读取model上下文参数
 * @author Vico
 * @create 2023-08-17 15:36
 */
@Data
@Accessors(chain = true)
@Builder
@AllArgsConstructor
@NoArgsConstructor
public class ExcelReaderModelContext<T extends ExcelBaseModel> {
    /**
     * model class
     */
    private Class<T> modelCla;
    /**
     * 单个model
     */
    private T model;
    /**
     * model list
     */
    private List<T> modelList;
    /**
     * 单个model 校验情况使用到
     */
    private ReadModelTitleConfig<T> listTitleConfig;
    /**
     * 解析list用到，记录行号
     */
    private Integer rowNum;
    /**
     * 导入缓存model
     */
    private ExcelCacheImportModel cacheImportModel;
    /**
     * model 参数
     */
    private ExcelReaderDataParam<T> param;
}
