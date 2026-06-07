package com.github.excel.model;

import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;
import lombok.experimental.Accessors;
import org.apache.poi.ss.util.CellRangeAddress;

/**
 * @author Vico
 * @create 2023-09-08 15:53
 */
@Data
@Accessors(chain = true)
@Builder
@AllArgsConstructor
@NoArgsConstructor
public class ExcelWriterMergeTitleModel {
    /**
     * 坐标
     */
    private ExcelWriterPointModel point;
    /**
     * 合并区域
     */
    private CellRangeAddress rangeAddress ;

}
