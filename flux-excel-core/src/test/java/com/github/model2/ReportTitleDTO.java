package com.github.model2;

import com.github.excel.annotation.ExcelWrite;
import com.github.excel.annotation.ExcelWriteProperty;
import com.github.excel.enums.ExcelWriterFillStyleEnum;
import com.github.excel.enums.ExcelThemeEnum;
import com.github.excel.model.ExcelBaseModel;
import lombok.Data;
import lombok.EqualsAndHashCode;

import java.util.List;

/**
 * 报告标题dto
 * @author Vico
 * @create 2022-08-19 15:14
 */
@Data
@ExcelWrite(themeName = ExcelThemeEnum.ZEBRA,freezeTitle = true)
@EqualsAndHashCode
public class ReportTitleDTO extends ExcelBaseModel {
    @ExcelWriteProperty(titleName = "id", disable = true, separator = "：", colWidth = 130, rowHeight = 30)
    private Long id;

    @ExcelWriteProperty(titleName = "所属企业", separator = "：", colWidth = 200, rowHeight = 30)
    private String groupName ;

    @ExcelWriteProperty(titleName = "适用企业", separator = "：", colWidth = 130, rowHeight = 30)
    private String applyEnterprise ;

    @ExcelWriteProperty(titleName = "参控股类型", separator = "：",fillStyle = ExcelWriterFillStyleEnum.VERTICAL, colWidth = 130, rowHeight = 30)
    private String shareType ;

    @ExcelWriteProperty(titleName = "参控股企业名称", separator = "：", colWidth = 200, rowHeight = 30)
    private String enterpriseName ;

    @ExcelWriteProperty(titleName = "数据周期", separator = "：",fillStyle = ExcelWriterFillStyleEnum.VERTICAL, colWidth = 200, rowHeight = 30)
    private String dataCycle ;

    @ExcelWriteProperty(titleName = "导出excel时用到",disable = true, separator = "：", colWidth = 130, rowHeight = 30)
    private List<Long> ids ;


}
