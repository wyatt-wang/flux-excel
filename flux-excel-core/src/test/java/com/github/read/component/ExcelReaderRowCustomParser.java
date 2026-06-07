package com.github.read.component;

import com.github.excel.context.ExcelReaderContext;
import com.github.excel.model.ExcelBaseModel;
import com.github.excel.read.handler.row.ExcelReaderRowParser;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;

/**
 * @author Vico
 * @create 2023-08-16 14:20
 */
@Slf4j
public class ExcelReaderRowCustomParser<T extends ExcelBaseModel> implements ExcelReaderRowParser<T> {
    DataFormatter df = new DataFormatter();
    @Override
    public void rowParser(ExcelReaderContext<T> readerContext) {
        for (Cell cell : readerContext.getParserContext().getRow()) {
            log.info(df.formatCellValue(cell));
        }
    }
}
