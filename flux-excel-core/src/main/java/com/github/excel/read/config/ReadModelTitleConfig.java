package com.github.excel.read.config;

import com.github.excel.model.ExcelBaseModel;
import lombok.Data;

import java.util.Map;
import java.util.concurrent.ConcurrentHashMap;

/**
 * @author Vico
 * @create 2024-01-27 11:10
 */
@Data
public class ReadModelTitleConfig<T extends ExcelBaseModel>{
    /**
     * 需要填充的model
     */
    Class<T> clazz;
    /**
     * 启始行，记录title的值，判断大于startRow
     */
    private Integer startRow ;
    /**
     * 结束列，用于判断行解析是否结束，执行回调
     */
    private Integer lastCol ;
    /**
     * 列作为key，value作为字段
     */
    private Map<Integer , ReadModelListTitleFieldConfig> fieldConfigMap = new ConcurrentHashMap<>();
    /**
     * 行作为key，bean实例作为value
     */
    private Map<Integer, T> rowBeanMap = new ConcurrentHashMap<>();

}


