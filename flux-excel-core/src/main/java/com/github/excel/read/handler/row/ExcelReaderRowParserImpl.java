package com.github.excel.read.handler.row;

import com.github.excel.context.ExcelReaderContext;
import com.github.excel.helper.ExcelHelper;
import com.github.excel.model.ExcelBaseModel;
import com.github.excel.read.handler.cell.ExcelReaderCellParser;
import com.github.excel.read.handler.cell.ExcelReaderCellParserImpl;
import com.github.excel.util.StringUtil;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;

import java.util.Map;
import java.util.Objects;

/**
 * Excel row标准解析
 * @author Vico
 * @create 2022-12-14 17:38
 */
public class ExcelReaderRowParserImpl<T extends ExcelBaseModel> extends AbstractExcelReaderRowParser<T> {

    private ExcelReaderCellParser<T> cellParser = new ExcelReaderCellParserImpl<T>();
    @Override
    public void beforeParser(ExcelReaderContext<T> readerContext) {

    }

    @Override
    public void afterParser(ExcelReaderContext<T> readerContext) {

    }

    @Override
    public void doParser(ExcelReaderContext<T> readerContext) {
        // 5.循环行对应所有的列
        int lastCellNum = readerContext.getParserContext().getRow().getLastCellNum();
        if (lastCellNum < 0) {
            return;
        }
        for (int cellIndex = 0; cellIndex < lastCellNum; cellIndex++) {
            Cell cell = readerContext.getParserContext().getRow().getCell(cellIndex, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
            if (Objects.isNull(cell)) {
                continue;
            }
            Object cellValue = ExcelHelper.getCellValue(cell, String.class, null, readerContext.getParserContext().getReaderFormatManager().getFormulaEvaluator());
            if (Boolean.TRUE.equals(readerContext.getReaderParam().getTrimString()) && cellValue instanceof String) {
                cellValue = ((String) cellValue).trim();
            }
            Object mergedCellValue = getMergedCellValue(readerContext, cell);
            if (StringUtil.isEmpty(cellValue) && Objects.nonNull(mergedCellValue)) {
                cellValue = mergedCellValue;
                if (cell.getCellType() == CellType.BLANK) {
                    cell.setCellValue(String.valueOf(mergedCellValue));
                }
            }
            readerContext.getParserContext().setCell(cell);
            readerContext.getParserContext().setCellValue(cellValue);
            cellParser.cellParser(readerContext);
        }
    }

    private Object getMergedCellValue(ExcelReaderContext<T> readerContext, Cell cell) {
        Map<String, Object> mergedCellValueMap = readerContext.getParserContext().getMergedCellValueMap();
        if (mergedCellValueMap == null || mergedCellValueMap.isEmpty()) {
            return null;
        }
        return mergedCellValueMap.get(cell.getRowIndex() + ":" + cell.getColumnIndex());
    }
}
