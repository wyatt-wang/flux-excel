package com.github.excel.param;

import com.github.excel.model.ExcelBaseModel;
import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;
import lombok.experimental.SuperBuilder;

import jakarta.validation.constraints.NotNull;
import java.util.List;

/**
 * Excel 写入list参数
 * @author Vico
 * @create 2023-08-17 15:36
 */
@Data
@AllArgsConstructor
@NoArgsConstructor
@SuperBuilder
public class ExcelWriterModelParam<T extends ExcelBaseModel> extends ExcelWriterDataParam {
    @NotNull(message = "modelCla不能为空")
    private Class<T> modelCla;

    @NotNull(message = "model不能为空")
    private T model;

    private List<String> excludeFields;
}
