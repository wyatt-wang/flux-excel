package com.github.excel.param;

import com.github.excel.validator.ExcelValidationGroup;
import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;
import lombok.experimental.Accessors;
import lombok.experimental.SuperBuilder;
import org.hibernate.validator.constraints.Range;

import jakarta.validation.constraints.NotEmpty;
import jakarta.validation.constraints.NotNull;

/**
 * Excel 写入参数
 * @author Vico
 * @create 2023-08-17 15:36
 */
@Data
@Accessors(chain = true)
@SuperBuilder
@AllArgsConstructor
@NoArgsConstructor
public abstract class ExcelWriterDataParam {
    @Builder.Default
    @NotNull(message = "行号不能为空")
    @Range(min = 0, max = 1048576 , message = "行号不能小于0且在xlsx下不能大于1048576" ,groups = {ExcelValidationGroup.XssfWorkbook.class})
    @Range(min = 0, max = 1048576 , message = "行号不能小于0且在xlsx下不能大于1048576" ,groups = {ExcelValidationGroup.SxssfWorkbook.class})
    @Range(min = 0, max = 65536 , message = "行号不能小于0且在xls下不能大于65536" ,groups = {ExcelValidationGroup.HssfWorkbook.class})
    private Integer rowIndex = 0;
    @Builder.Default
    @NotNull(message = "列号不能为空")
    @Range(min = 0, max = 16384, message = "列号不能小于0且在xlsx下不能大于16384", groups = {ExcelValidationGroup.XssfWorkbook.class})
    @Range(min = 0, max = 16384, message = "列号不能小于0且在xlsx下不能大于16384", groups = {ExcelValidationGroup.SxssfWorkbook.class})
    @Range(min = 0, max = 255, message = "列号不能小于0且在xls下不能大于255", groups = {ExcelValidationGroup.HssfWorkbook.class})
    private Integer colIndex = 0;

    @NotEmpty(message = "sheetName不能为空")
    private String sheetName;

    @NotNull(message = "fillTemplate不能为空")
    @Builder.Default
    private Boolean fillTemplate = true;
    /**
     * 校验
     */
    private ExcelWriterValidationParam validationParam ;
    /**
     * 批注
     */
    private ExcelWriterCommentParam commentParam ;
}
