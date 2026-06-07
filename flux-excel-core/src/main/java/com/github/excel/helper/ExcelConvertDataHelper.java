package com.github.excel.helper;

import com.github.excel.constant.ExcelConstant;
import com.github.excel.exception.ExcelReaderException;
import com.github.excel.model.ExcelCacheImportModel;
import com.github.excel.util.StringUtil;
import com.google.common.base.Throwables;
import lombok.extern.slf4j.Slf4j;

import java.math.BigDecimal;
import java.text.SimpleDateFormat;
import java.util.Date;

/**
 * 转换数据帮助类
 * @author Vico
 * @create 2022-12-14 19:33
 */
@Slf4j
public class ExcelConvertDataHelper {
    public static Object convertDataType(Object setParams, ExcelCacheImportModel.ExcelCacheImportFieldModel fieldModel) {
        if (null == setParams) {
            return null;
        }
        try {
            if(setParams.getClass().equals(Double.class) && Long.class.equals(fieldModel.getSetMethod().getParameterTypes()[0])){
                setParams = ((Double)setParams).longValue();
            }else if(setParams.getClass().equals(Double.class) && Integer.class.equals(fieldModel.getSetMethod().getParameterTypes()[0])){
                setParams = ((Double)setParams).intValue();
            }else if(setParams.getClass().equals(Double.class) && Float.class.equals(fieldModel.getSetMethod().getParameterTypes()[0])){
                setParams = ((Double)setParams).floatValue();
            }else if(setParams.getClass().equals(Double.class) && Short.class.equals(fieldModel.getSetMethod().getParameterTypes()[0])){
                setParams = ((Double)setParams).shortValue();
            }else if(setParams.getClass().equals(Double.class) && Byte.class.equals(fieldModel.getSetMethod().getParameterTypes()[0])){
                setParams = ((Double)setParams).byteValue();
            }else if(setParams.getClass().equals(String.class) && Double.class.equals(fieldModel.getSetMethod().getParameterTypes()[0])){
                setParams = new BigDecimal(setParams.toString()).doubleValue();
            }else if(setParams.getClass().equals(String.class) && Float.class.equals(fieldModel.getSetMethod().getParameterTypes()[0])){
                setParams = new BigDecimal(setParams.toString()).floatValue();
            }else if(setParams.getClass().equals(String.class) && Long.class.equals(fieldModel.getSetMethod().getParameterTypes()[0])){
                setParams = new BigDecimal(setParams.toString()).longValue();
            }else if(setParams.getClass().equals(String.class) && Integer.class.equals(fieldModel.getSetMethod().getParameterTypes()[0])){
                setParams = new BigDecimal(setParams.toString()).intValue();
            }else if(setParams.getClass().equals(String.class) && Short.class.equals(fieldModel.getSetMethod().getParameterTypes()[0])){
                setParams = new BigDecimal(setParams.toString()).shortValue();
            }else if(setParams.getClass().equals(String.class) && Byte.class.equals(fieldModel.getSetMethod().getParameterTypes()[0])){
                setParams = new BigDecimal(setParams.toString()).byteValue();
            }else if(setParams.getClass().equals(Date.class) && String.class.equals(fieldModel.getSetMethod().getParameterTypes()[0])){
                String format = fieldModel.getImportProperty().formatPattern();
                if (StringUtil.isEmpty(format)) {
                    format = ExcelConstant.DEFAULT_DATE_FORMAT;
                }
                SimpleDateFormat sdf = new SimpleDateFormat(format);
                setParams = sdf.format((Date) setParams);
            }else if(setParams.getClass().equals(Double.class) && String.class.equals(fieldModel.getSetMethod().getParameterTypes()[0])){
                setParams = String.valueOf(setParams);
            }
        } catch (Exception e) {
            log.error(Throwables.getStackTraceAsString(e));
            throw new ExcelReaderException("转换失败，接收字段"+fieldModel.getField()+"，接收类型：" + fieldModel.getSetMethod().getParameterTypes()[0].getTypeName() + "，转换值：" + setParams);
        }
        return setParams;
    }
}
