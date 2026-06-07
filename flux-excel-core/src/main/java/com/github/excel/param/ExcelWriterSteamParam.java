package com.github.excel.param;

import lombok.Data;
import lombok.EqualsAndHashCode;
import lombok.ToString;
import lombok.experimental.Accessors;
import lombok.experimental.SuperBuilder;

import jakarta.validation.constraints.NotNull;
import java.io.OutputStream;

/**
 * Excel 文件流写入参数
 * @author Vico
 * @create 2023-08-17 15:36
 */
@Data
@Accessors(chain = true)
@SuperBuilder
@ToString
@EqualsAndHashCode
public class ExcelWriterSteamParam extends ExcelWriterParam{
    /**
     * 流
     */
    @NotNull(message = "outputStream不能为空")
    private OutputStream outputStream;
}
