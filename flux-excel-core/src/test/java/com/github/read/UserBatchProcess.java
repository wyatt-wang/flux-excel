package com.github.read;

import com.github.excel.read.facade.ExcelReaderBatchProcess;
import com.github.model.UserExcelDtoImportBean;
import lombok.extern.slf4j.Slf4j;

import java.util.List;

/**
 * @author Vico
 * @create 2022-12-16 15:02
 */
@Slf4j
public class UserBatchProcess implements ExcelReaderBatchProcess<UserExcelDtoImportBean> {
    @Override
    public int getBatchSize() {
        return 2;
    }

    @Override
    public void process(List<UserExcelDtoImportBean> dataList) {
        dataList.forEach(e->log.info(e.toString()));
    }
}
