package com.github.excel.read.handler.row;

import com.github.excel.context.ExcelReaderContext;
import com.github.excel.helper.ExcelHelper;
import com.github.excel.model.ExcelBaseModel;
import com.github.excel.read.handler.cell.ExcelReaderCellParser;
import com.github.excel.read.handler.cell.ExcelReaderCellParserImpl;
import org.apache.poi.ss.usermodel.Cell;

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
        for (int cellIndex = readerContext.getParserContext().getRow().getFirstCellNum(); cellIndex < readerContext.getParserContext().getRow().getLastCellNum(); cellIndex++) {
            Cell cell = readerContext.getParserContext().getRow().getCell(cellIndex);
            if (Objects.isNull(cell)) {
                continue;
            }
            Object cellValue = ExcelHelper.getCellValue(cell, String.class, null, readerContext.getParserContext().getReaderFormatManager().getFormulaEvaluator());
            if (Boolean.TRUE.equals(readerContext.getReaderParam().getTrimString()) && cellValue instanceof String) {
                cellValue = ((String) cellValue).trim();
            }
            readerContext.getParserContext().setCell(cell);
            readerContext.getParserContext().setCellValue(cellValue);
            cellParser.cellParser(readerContext);
        }
    }
}
