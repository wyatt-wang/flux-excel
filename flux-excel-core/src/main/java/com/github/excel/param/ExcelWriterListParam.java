package com.github.excel.param;

import com.github.excel.model.ExcelBaseModel;
import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;
import lombok.experimental.Accessors;
import lombok.experimental.SuperBuilder;

import jakarta.validation.constraints.NotNull;
import jakarta.validation.constraints.Size;
import java.util.List;

/**
 * Excel 写入list参数
 * @author Vico
 * @create 2023-08-17 15:36
 */
@Data
@Accessors(chain = true)
@AllArgsConstructor
@NoArgsConstructor
@SuperBuilder
public class ExcelWriterListParam<T extends ExcelBaseModel> extends ExcelWriterDataParam {
    @NotNull(message = "modelCla不能为空")
    private Class<T> modelCla;

    @NotNull(message = "modelList不能为空")
    @Size(min = 1 , message = "modelList最少有1个元素")
    private List<T> modelList;
    /**
     * 排除字段
     */
    private List<String> excludeFields;
}
