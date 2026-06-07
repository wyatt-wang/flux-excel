package com.github.excel.param;

import com.github.excel.constant.ExcelConstant;
import com.github.excel.enums.ExcelSuffixEnum;
import lombok.Builder;
import lombok.Data;
import lombok.EqualsAndHashCode;
import lombok.ToString;
import lombok.experimental.Accessors;
import lombok.experimental.SuperBuilder;

import jakarta.validation.constraints.NotNull;

/**
 * Excel 写入参数
 * @author Vico
 * @create 2023-08-17 15:36
 */
@Data
@Accessors(chain = true)
@SuperBuilder
@ToString
@EqualsAndHashCode
public abstract class ExcelWriterParam {
    /**
     * 流大小
     */
    @Builder.Default
    private Integer rowAccessWindowSize = ExcelConstant.DEFAULT_ROW_ACCESS_WINDOW_SIZE;
    /**
     * SXSSF 临时文件是否压缩。
     */
    @Builder.Default
    private Boolean compressTempFiles = false;
    /**
     * SXSSF 是否使用共享字符串表。
     */
    @Builder.Default
    private Boolean useSharedStringsTable = false;
    /**
     * 是否启用流模式导出
     */
    @Builder.Default
    @NotNull(message = "streaming不能为空")
    private Boolean streaming = false;
    /**
     * 选择sheet
     */
    private String selectSheetName;
    /**
     * 填充没有数据提示
     */
    @Builder.Default
    @NotNull(message = "noneDataTips不能为空")
    private Boolean noneDataTips = true;
    /**
     * 没有数据提示信息
     */
    private String noneDataTipsMsg;
    /**
     * sheet 最大行数
     */
    private Integer sheetRowMaxCount ;
    /**
     * 模版名称
     */
    private String template;
    /**
     * 模版磁盘路径
     */
    private String templateFilePath;
    /**
     * 模版密码
     */
    private String templatePassword;
    /**
     * 后缀名称
     */
    @NotNull(message = "suffixEnum不能为空")
    private ExcelSuffixEnum suffixEnum;
}
