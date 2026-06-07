package com.github.excel.read.handler.row;

import com.github.excel.context.ExcelReaderContext;
import com.github.excel.model.ExcelBaseModel;

/**
 * Excel row解析
 * @author Vico
 * @create 2022-12-14 17:34
 */
public interface ExcelReaderRowParser<T extends ExcelBaseModel> {
    void rowParser(ExcelReaderContext<T> readerContext);
}
