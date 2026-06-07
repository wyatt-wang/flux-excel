package com.github.excel.param;

import com.github.excel.constant.ExcelConstant;
import com.github.excel.enums.ExcelWriterColumnFillTypeEnum;
import com.github.excel.write.ExcelDefaultWriterDataFormat;
import com.github.excel.write.ExcelWriterDataFormat;
import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;
import lombok.experimental.Accessors;
import lombok.experimental.SuperBuilder;

/**
 * @Author: Vachel Wang
 * @Date: 2026/4/24
 * @Description: 导出到单个单元格model
 */
@Data
@Accessors(chain = true)
@SuperBuilder
@AllArgsConstructor
@NoArgsConstructor
public class ExcelWriterCellParam extends ExcelWriterDataParam{
    /**
     * 行高
     */
    private Short rowHeight = ExcelConstant.MINUS_TWO_SHORT;
    /**
     * 列宽
     */
    private Short colWidth = ExcelConstant.MINUS_TWO_SHORT;
    /**
     * 样式名称
     */
    private String styleName;
    /**
     * 值
     */
    private Object value;
    /**
     * 格式化字符串
     */
    private String formatPattern;
    /**
     * 数据格式化
     */
    private Class<? extends ExcelWriterDataFormat> dataFormat = ExcelDefaultWriterDataFormat.class;
    /**
     * 填充类型
     */
    private ExcelWriterColumnFillTypeEnum fillType = ExcelWriterColumnFillTypeEnum.COVER;
}
