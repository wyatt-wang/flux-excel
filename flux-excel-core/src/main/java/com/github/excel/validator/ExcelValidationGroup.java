package com.github.excel.validator;

/**
 * 校验分组
 * @author Vico
 * @create 2024-05-12 12:02
 */
public class ExcelValidationGroup {
    /**
     * xls 分组
     */
    public interface HssfWorkbook{}

    /**
     * xlsx 分组
     */
    public interface XssfWorkbook{}

    /**
     * xlsx stream 分组
     */
    public interface SxssfWorkbook{}
}
