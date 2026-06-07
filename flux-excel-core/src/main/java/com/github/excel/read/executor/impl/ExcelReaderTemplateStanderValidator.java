package com.github.excel.read.executor.impl;

import com.github.excel.boot.ExcelBootLoader;
import com.github.excel.constant.ExcelConstant;
import com.github.excel.constant.ExcelErrorMsgConstant;
import com.github.excel.exception.ExcelReaderException;
import com.github.excel.helper.ExcelHelper;
import com.github.excel.model.ExcelImportTemplateCacheModel;
import com.github.excel.read.executor.ExcelReaderTemplateValidator;
import com.github.excel.util.StringUtil;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.util.List;
import java.util.Map;
import java.util.Objects;

/**
 * Excel 标准校验器
 * @author Vico
 * @create 2022-12-14 16:42
 */
public class ExcelReaderTemplateStanderValidator implements ExcelReaderTemplateValidator {

    /**
     * 校验模版
     * @param template
     * @param workbook
     */
    @Override
    public void validateTemplate(String template, Workbook workbook) {
        Map<String, List<ExcelImportTemplateCacheModel>> templateCacheMap = ExcelBootLoader.getExcelImportTemplateCacheMapValue(template);
        if (Objects.nonNull(templateCacheMap)) {
            for (Map.Entry<String, List<ExcelImportTemplateCacheModel>> entry : templateCacheMap.entrySet()) {
                Sheet sheet = workbook.getSheet(entry.getKey());
                if (Objects.isNull(sheet)) {
                    throw new ExcelReaderException(String.format(ExcelErrorMsgConstant.ERROR_NOT_FOUND_SHEET, entry.getKey()));
                }
                List<ExcelImportTemplateCacheModel> cacheModelList = entry.getValue();
                for (ExcelImportTemplateCacheModel cacheModel : cacheModelList) {
                    Row row = sheet.getRow(cacheModel.getRowIndex());
                    if (Objects.isNull(row)) {
                        throw new ExcelReaderException(String.format(ExcelErrorMsgConstant.ERROR_NOT_FOUND_SHEET_ROW, entry.getKey(), String.valueOf(cacheModel.getRowIndex() + ExcelConstant.ONE_INT)));
                    }
                    Cell cell = row.getCell(cacheModel.getColIndex());
                    if (Objects.isNull(cell)) {
                        throw new ExcelReaderException(String.format(ExcelErrorMsgConstant.ERROR_NOT_FOUND_SHEET_ROW_COL, entry.getKey(), String.valueOf(cacheModel.getRowIndex() + ExcelConstant.ONE_INT), String.valueOf(cacheModel.getColIndex() + ExcelConstant.ONE_INT)));
                    }
                    Object value = ExcelHelper.getCellValue(cell, String.class, null,workbook.getCreationHelper().createFormulaEvaluator());
                    if (StringUtil.isEmpty(value) || !value.equals(cacheModel.getText())) {
                        String errorTips = String.format(ExcelErrorMsgConstant.ERROR_NOT_MATCH_COL_CONTENT, entry.getKey(), cell.getAddress().formatAsString(), cacheModel.getText());
                        throw new ExcelReaderException(errorTips);
                    }
                }
            }
        }
    }
}
