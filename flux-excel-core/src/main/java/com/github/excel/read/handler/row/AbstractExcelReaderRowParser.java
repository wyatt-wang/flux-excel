package com.github.excel.read.handler.row;

import com.github.excel.context.ExcelReaderContext;
import com.github.excel.model.ExcelBaseModel;

/**
 * @Author: Vachel Wang
 * @Date: 2026/4/24
 * @Description: 批量处理大小
 */
public abstract class AbstractExcelReaderRowParser<T extends ExcelBaseModel> implements ExcelReaderRowParser<T> {

    public abstract void beforeParser(ExcelReaderContext<T> readerContext);

    public abstract void afterParser(ExcelReaderContext<T> readerContext);

    public abstract void doParser(ExcelReaderContext<T> readerContext);

    public void rowParser(ExcelReaderContext<T> readerContext){
        beforeParser(readerContext);
        doParser(readerContext);
        afterParser(readerContext);
    }

}
