package com.github.excel.model;

import com.github.excel.annotation.ExcelWriteProperty;
import lombok.Builder;
import lombok.Data;

import java.lang.reflect.Method;

/**
 * @author Vico
 * @create 2024-02-18 14:44
 */
@Data
@Builder
public class ExcelWriterTemplateModel {
    /**
     * 方法
     */
    private Method getMethod;
    /**
     * 行号
     */
    private Integer rowIndex;
    /**
     * 列号
     */
    private Integer colIndex;
    /**
     * sheet名称
     */
    private String sheetName;
    /**
     * 缓存字段
     */
    private ExcelCacheFieldModel cacheFieldModel;
    /**
     * 单元格注解
     */
    private ExcelWriteProperty exportCell;
}
