package com.github.excel.param;

import com.github.excel.model.ExcelReadError;
import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;
import lombok.experimental.Accessors;
import lombok.experimental.SuperBuilder;

import jakarta.validation.constraints.NotNull;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.function.Function;

/**
 * Excel 读取参数
 * @author Vico
 * @create 2023-08-17 15:36
 */
@Data
@Accessors(chain = true)
@SuperBuilder
@AllArgsConstructor
@NoArgsConstructor
public abstract class ExcelReaderParam{
    /**
     * 是否读取图片
     */
    @Builder.Default
    @NotNull(message = "readPicture 不能为空")
    private Boolean readPicture = false;
    /**
     * 是否关闭输入流
     */
    @Builder.Default
    @NotNull(message = "closeInputStream 不能为空")
    private Boolean closeInputStream = false;
    /**
     * 是否只读
     */
    @Builder.Default
    @NotNull(message = "readOnly 不能为空")
    private Boolean readOnly = false;
    /**
     * 密码
     */
    private String password;

    /**
     * 模版
     */
    private String template;
    /**
     * 是否忽略空行
     */
    @Builder.Default
    private Boolean ignoreEmptyRow = true;
    /**
     * 是否裁剪字符串空白
     */
    @Builder.Default
    private Boolean trimString = false;
    /**
     * 合并单元格策略
     */
    private String mergedCellStrategy;
    /**
     * 空单元格策略
     */
    private String emptyCellPolicy;
    /**
     * 是否收集单元格错误并继续解析
     */
    @Builder.Default
    private Boolean collectErrors = false;
    /**
     * 单元格错误明细
     */
    @Builder.Default
    private List<ExcelReadError> readErrors = new ArrayList<>();
    /**
     * 导入类型转换器
     */
    @Builder.Default
    private Map<Class<?>, Function<Object, Object>> typeConverters = new LinkedHashMap<>();
    /**
     * 导入字段转换器
     */
    @Builder.Default
    private Map<String, Function<Object, Object>> fieldConverters = new LinkedHashMap<>();

}
