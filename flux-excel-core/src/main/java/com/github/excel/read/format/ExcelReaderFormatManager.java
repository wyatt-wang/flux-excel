package com.github.excel.read.format;

import com.github.excel.exception.ExcelReaderException;
import com.google.common.base.Throwables;
import com.google.common.collect.Maps;
import lombok.Getter;
import lombok.Setter;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.FormulaEvaluator;

import java.util.Map;

/**
 * @author Vico
 * @create 2022-12-14 19:48
 */
@Slf4j
public class ExcelReaderFormatManager {
    protected Map<Class<? extends ExcelReaderDataFormat>, ExcelReaderDataFormat> dataFormatMap = Maps.newHashMap();
    protected ExcelReaderDataFormat defaultFormat = new ExcelDefaultReaderDataFormat();

    @Getter
    @Setter
    private FormulaEvaluator formulaEvaluator;

    public ExcelReaderDataFormat getDataFormatThenCache(Class<? extends ExcelReaderDataFormat> formatCla) {
        if (formatCla == ExcelDefaultReaderDataFormat.class) {
            return defaultFormat;
        } else {
            ExcelReaderDataFormat excelReaderDataFormat = dataFormatMap.get(formatCla);
            if (excelReaderDataFormat != null) {
                return excelReaderDataFormat;
            }
            try {
                excelReaderDataFormat = formatCla.newInstance();
            } catch (InstantiationException e) {
                log.error("Read excel failed , cause:{}", Throwables.getStackTraceAsString(e));
                throw new ExcelReaderException(e.getMessage());
            } catch (IllegalAccessException e) {
                log.error("Read excel failed , cause:{}", Throwables.getStackTraceAsString(e));
                throw new ExcelReaderException(e.getMessage());
            }
            dataFormatMap.put(formatCla, excelReaderDataFormat);
            return excelReaderDataFormat;
        }

    }

}
