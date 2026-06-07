package com.github.excel.read.config;

import com.github.excel.model.ExcelCacheImportModel;
import lombok.AllArgsConstructor;
import lombok.Data;

/**
 * @author Vico
 * @create 2024-01-27 11:14
 */
@Data
@AllArgsConstructor
public class ReadModelListTitleFieldConfig{
    /**
     * 需要填充的model field信息
     */
    private ExcelCacheImportModel.ExcelCacheImportFieldModel cacheImportFieldModel ;
    /**
     * 标题名称
     */
    private String titleName ;
    /**
     * 运行时行行号
     */
    private Integer rowIndex ;
}
