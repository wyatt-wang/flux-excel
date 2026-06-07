package com.github.excel.param;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;
import lombok.experimental.Accessors;
import lombok.experimental.SuperBuilder;

import jakarta.validation.constraints.NotNull;
import java.io.File;

/**
 * Excel file 读取参数
 * @author Vico
 * @create 2023-08-17 15:36
 */
@Data
@Accessors(chain = true)
@SuperBuilder
@AllArgsConstructor
@NoArgsConstructor
public class ExcelReaderFileParam extends ExcelReaderParam{

    @NotNull(message = "文件不能为空")
    private File file;

}
