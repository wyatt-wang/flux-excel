package com.github.excel.model;

import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.experimental.Accessors;

/**
 * @author Vico
 * @create 2023-09-08 15:53
 */
@Data
@Accessors(chain = true)
@Builder
@AllArgsConstructor
public class ExcelWriterPointModel {
    /**
     * 开始列下标
     */
    private final int firstColIndex;
    /**
     * 列下标，此参数将动态变更
     */
    private int colIndex;
    /**
     * 最后列index
     */
    private int lastColIndex;
    /**
     * 开始行下标
     */
    private final int firstRowIndex;
    /**
     * 行下标，此参数将动态变更
     */
    private  int rowIndex;
    /**
     * 最后行index
     */
    private int lastRowIndex;

}
