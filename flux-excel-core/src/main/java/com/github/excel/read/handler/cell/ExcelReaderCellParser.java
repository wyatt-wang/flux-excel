package com.github.excel.read.handler.cell;

import com.github.excel.context.ExcelReaderContext;
import com.github.excel.model.ExcelBaseModel;

/**
 * Excel 单元格解析
 * @author Vico
 * @create 2022-12-14 17:34
 */
public interface ExcelReaderCellParser<T extends ExcelBaseModel> {
    void cellParser(ExcelReaderContext<T> readerContext);
}
