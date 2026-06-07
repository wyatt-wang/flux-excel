package com.github.excel.read.listener;

import lombok.Data;
import lombok.experimental.Accessors;

/**
 * Excel 导入监听器上下文
 */
@Data
@Accessors(chain = true)
public class ExcelReadListenerContext {
	private Integer sheetIndex;
	private String sheetName;
	private Integer rowIndex;
}
