package com.github.model;

/**
 * @author Vico
 * @create 2022-08-18 14:55
 */

import com.github.excel.annotation.ExcelWrite;
import com.github.excel.annotation.ExcelWriteProperty;
import com.github.excel.enums.ExcelWriterFillStyleEnum;
import com.github.excel.enums.ExcelThemeEnum;
import com.github.excel.model.ExcelBaseModel;
import lombok.Builder;
import lombok.Data;
import lombok.EqualsAndHashCode;

/**
 * @ClassName: Accessory1DTO
 * @Description: 附件1数据传输DTO
 * @Author: Vachel Wang
 * @Date: 2026/4/24
 **/
@Data
@Builder
@EqualsAndHashCode(callSuper = true)
@ExcelWrite(themeName = ExcelThemeEnum.ZEBRA, fillContent = false)
public class Accessory1DTO extends ExcelBaseModel {

    /**
     * 此行数据在set时不需要关注, 是标题名称
     */
    @ExcelWriteProperty(titleName = "地方国资委及所属国有企业控股金融子企业基本情况表",mergeRowNum = 1)
    private String titleName;

    @ExcelWriteProperty(titleName = "行次",mergeRowNum = 1)
    private String sequenceNo;

    @ExcelWriteProperty(titleName = "集团标准全称",mergeRowNum = 1)
    private String groupName;

    @ExcelWriteProperty(titleName = "金融子企业名称",mergeTitleColNum = 2)
    private String enterpriseName;

    @ExcelWriteProperty(titleName = "金融行业类型-国务院国资委",mergeTitleColNum = 2)
    private String financialIndustry;

    @ExcelWriteProperty(titleName = "金融机构二级分类",verticalNewLine = true,colIndex = 3,fillStyle = ExcelWriterFillStyleEnum.VERTICAL)
    private String financeSecType;

    @ExcelWriteProperty(titleName = "修改原因",colIndex = 4)
    private String modReason;

    @ExcelWriteProperty(titleName = "合计持股比例",colIndex = 5)
    private String shareHoldingRat;

    @ExcelWriteProperty(titleName = "主要股东标准全称及持股比例",colIndex = 6)
    private String largestShareholderNameAndRat;

    @ExcelWriteProperty(titleName = "管理层级",colIndex = 7)
    private String managementLevel;

    @ExcelWriteProperty(titleName = "股权层级",colIndex = 8)
    private String equityLevel;


}
