package com.github.excel.param;

import com.github.excel.model.ExcelBaseModel;
import com.github.excel.read.handler.row.ExcelReaderRowHandler;
import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;
import lombok.experimental.Accessors;
import lombok.experimental.SuperBuilder;
import org.hibernate.validator.constraints.Range;

import jakarta.validation.constraints.NotNull;

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
public abstract class ExcelReaderDataParam<T extends ExcelBaseModel> {
    /**
     * sheet 下标
     */
    @Range(min = 0,max = 254,message = "sheetIndex 支持范围为[0-254]")
    @NotNull(message = "sheetIndex 不能为空")
    private Integer sheetIndex = 0;
    /**
     * sheet 名称，优先级高于 sheetIndex
     */
    private String sheetName;
    /**
     * 表头行数
     */
    private Integer headRowNumber;
    /**
     * 数据起始行
     */
    private Integer dataStartRow;
    /**
     * 数据结束行
     */
    private Integer dataEndRow;
    /**
     * 最大读取数据行数
     */
    private Integer maxRows;
    /**
     * model class
     */
    @NotNull(message = "modelCla不能为空")
    private Class<T> modelCla;
    /**
     * 行处理器
     */
    private ExcelReaderRowHandler<T> rowHandler;

}
