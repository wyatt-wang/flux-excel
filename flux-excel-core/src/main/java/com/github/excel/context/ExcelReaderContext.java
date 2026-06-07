package com.github.excel.context;

import com.github.excel.model.ExcelBaseModel;
import com.github.excel.param.ExcelReaderParam;
import com.github.excel.read.facade.ExcelCustomReader;
import com.google.common.collect.Maps;
import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;
import lombok.experimental.Accessors;

import java.util.Map;

/**
 * Excel 读取上下文参数
 * @author Vico
 * @create 2023-08-17 15:36
 */
@Data
@Accessors(chain = true)
@Builder
@AllArgsConstructor
@NoArgsConstructor
public class ExcelReaderContext<T extends ExcelBaseModel> {
    /**
     * 读取参数
     */
    private ExcelReaderParam readerParam;
    /**
     * 自定义读
     */
    private ExcelCustomReader excelCustomReader;

    /**
     * 待解析 model 缓存
     */
    private final Map<Class<T>, ExcelReaderModelContext<T>> modelMap = Maps.newHashMap();
    /**
     * excel 解析上下文
     */
    private ExcelReaderParserContext<T> parserContext = new ExcelReaderParserContext<T>();
}
