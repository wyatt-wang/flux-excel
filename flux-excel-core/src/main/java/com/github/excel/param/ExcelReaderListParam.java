package com.github.excel.param;

import com.github.excel.model.ExcelBaseModel;
import com.github.excel.read.facade.ExcelReaderBatchProcess;
import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;
import lombok.experimental.Accessors;
import lombok.experimental.SuperBuilder;

/**
 * Excel 写入参数
 * @author Vico
 * @create 2023-08-17 15:36
 */
@Data
@Accessors(chain = true)
@SuperBuilder
@AllArgsConstructor
@NoArgsConstructor
public class ExcelReaderListParam<T extends ExcelBaseModel> extends ExcelReaderDataParam<T> {
    /**
     * Excel 读取批处理
     */
    private ExcelReaderBatchProcess<T> batchProcess;
}
