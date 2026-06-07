package com.github.excel.read.handler.row;

import com.github.excel.model.ExcelBaseModel;

/**
 * @Author: Vachel Wang
 * @Date: 2026/4/24
 * @Description: 自定义写读接口
 */
@FunctionalInterface
public interface ExcelReaderRowHandler<T extends ExcelBaseModel> {
	void handler(T model);
}
