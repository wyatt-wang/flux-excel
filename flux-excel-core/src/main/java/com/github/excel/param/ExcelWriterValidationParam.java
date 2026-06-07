package com.github.excel.param;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;
import lombok.experimental.Accessors;
import lombok.experimental.SuperBuilder;

/**
 * Excel 校验参数
 * @author Vico
 * @create 2023-08-17 15:36
 */
@Data
@Accessors(chain = true)
@SuperBuilder
@AllArgsConstructor
@NoArgsConstructor
public abstract class ExcelWriterValidationParam {
    /**
     * 当前单元格的原始值
     */
    private Object value;
    /**
     * 标题
     */
	private String title ;
    /**
     * 消息
     */
    private String message ;
}
