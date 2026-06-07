package com.github.excel.model;

import com.github.excel.annotation.ExcelWriteProperty;
import com.github.excel.constant.ExcelConstant;
import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;
import lombok.experimental.Accessors;

import java.lang.reflect.Field;
import java.lang.reflect.Method;

/**
 * @author Vico
 * @create 2023-09-14 11:04
 */
@Data
@NoArgsConstructor
@AllArgsConstructor
@Accessors(chain = true)
public class ExcelCacheFieldModel {
    /**
     * exportCell 注解
     */
    private ExcelWriteProperty exportCell;
    /**
     * get 函数
     */
    private Method getMethod ;
    /**
     * 动态超链接展示名 get 函数
     */
    private Method linkNameGetMethod;
    /**
     * 字段
     */
    private Field field ;
    /**
     * 字段名称
     */
    private String fieldName ;
    /**
     * 标题样式名称
     */
    private String titleStyleName ;
    /**
     * 内容样式名称
     */
    private String contentStyleName ;
    /**
     * 偶数行样式名称
     */
    private String evenRowStyleName ;
    /**
     * 标题默认高度
     */
    private Short titleRowHeight;
    /**
     * 内容默认高度
     */
    private Short contentRowHeight;
    /**
     * 列表高度
     */
    private Short listRowHeight;
    /**
     * 单元格默认宽度
     */
    private Short colWidth;
    /**
     * 是否map类型
     */
    private boolean isMap;
    /**
     * 排序号
     */
    private Integer index = ExcelConstant.INT_10000;
    /**
     * 标题名称，目前用于在生成序列号时使用
     */
    private String titleName ;
}
