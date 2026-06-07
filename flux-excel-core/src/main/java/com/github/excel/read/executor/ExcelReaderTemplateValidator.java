package com.github.excel.read.executor;

import org.apache.poi.ss.usermodel.Workbook;

/**
 * Excel 模版校验器
 * @author Vico
 * @create 2022-12-14 16:34
 */
public interface ExcelReaderTemplateValidator {
    /**
     * 校验模版
     * @param template 模版
     * @param workbook 文档
     */
    void validateTemplate(String template, Workbook workbook);

}
