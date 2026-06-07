package com.github.excel.param;

import com.github.excel.model.ExcelBaseModel;
import lombok.Data;
import lombok.EqualsAndHashCode;
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
@EqualsAndHashCode(callSuper = true)
public class ExcelReaderModelParam<T extends ExcelBaseModel> extends ExcelReaderDataParam<T> {

}
