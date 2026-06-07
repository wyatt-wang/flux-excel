package com.github.read.handler;

import com.github.excel.read.handler.row.ExcelReaderRowHandler;
import com.github.model.UserExcelDtoImportBean;
import lombok.extern.slf4j.Slf4j;

/**
 * @author Vico
 * @create 2023-08-16 16:49
 */
@Slf4j
public class ExcelStandardRowHandler implements ExcelReaderRowHandler<UserExcelDtoImportBean> {
    @Override
    public void handler(UserExcelDtoImportBean model) {
        log.info(model.toString());
    }
}
