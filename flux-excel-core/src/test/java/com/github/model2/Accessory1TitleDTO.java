package com.github.model2;

import com.github.excel.annotation.ExcelWrite;
import com.github.excel.annotation.ExcelWriteProperty;
import com.github.excel.enums.ExcelThemeEnum;
import com.github.excel.model.ExcelBaseModel;
import lombok.Data;

/**
 * @ClassName: Accessory1TitleDTO
 * @Description: 附件1模版表头DTO
 * @Author: Vachel Wang
 * @Date: 2026/4/24
 **/
@Data
@ExcelWrite(themeName = ExcelThemeEnum.ZEBRA, fillContent = false,freezeTitle = true)
public class Accessory1TitleDTO extends ExcelBaseModel {

    /**
     * 此行数据在set时不需要关注, 是标题名称
     */
    @ExcelWriteProperty(titleName = "地方国资委及所属国有企业控股金融子企业基本情况表", rowIndex = 2, colIndex = 0, mergeTitleColNum = 28)
    private String titleName;

    @ExcelWriteProperty(titleName = "行次", rowIndex = 3, colIndex = 0, mergeRowNum = 1)
    private String sequenceNo;

    @ExcelWriteProperty(titleName = "集团标准全称", rowIndex = 3, colIndex = 1, mergeRowNum = 1,
            colWidth = 200)
    private String groupName;

    @ExcelWriteProperty(titleName = "金融子企业名称", rowIndex = 3, colIndex = 2, mergeRowNum = 1,
            colWidth = 200)
    private String enterpriseName;

    @ExcelWriteProperty(titleName = "金融行业类型-国务院国资委", rowIndex = 3, colIndex = 3, mergeRowNum = 1,
            colWidth = 300)
    private String financialIndustry;

    @ExcelWriteProperty(titleName = "金融机构二级分类", rowIndex = 3, colIndex = 4, mergeRowNum = 1,
            colWidth = 300)
    private String financeSecType;

    @ExcelWriteProperty(titleName = "修改原因", rowIndex = 3, colIndex = 5, mergeRowNum = 1,
            colWidth = 100)
    private String modReason;

    @ExcelWriteProperty(titleName = "合计持股比例", rowIndex = 3, colIndex = 6, mergeRowNum = 1,
            colWidth = 200)
    private String shareHoldingRat;

    @ExcelWriteProperty(titleName = "主要股东标准全称及持股比例", rowIndex = 3, colIndex = 7, mergeRowNum = 1,
            colWidth = 300)
    private String largestShareholderNameAndRat;

    @ExcelWriteProperty(titleName = "管理层级", rowIndex = 3, colIndex = 8, mergeRowNum = 1,
            colWidth = 200)
    private String managementLevel;

    @ExcelWriteProperty(titleName = "股权层级", rowIndex = 3, colIndex = 9, mergeRowNum = 1,
            colWidth = 200)
    private String equityLevel;

    @ExcelWriteProperty(titleName = "总资产", rowIndex = 3, colIndex = 10, mergeRowNum = 1,
            colWidth = 200)
    private String totalAssets;

    @ExcelWriteProperty(titleName = "净资产", rowIndex = 3, colIndex = 11, mergeRowNum = 1,
            colWidth = 200)
    private String netAssets;

    @ExcelWriteProperty(titleName = "营业收入", rowIndex = 3, colIndex = 12, mergeRowNum = 1,
            colWidth = 200)
    private String income;

    @ExcelWriteProperty(titleName = "利润总额", rowIndex = 3, colIndex = 13, mergeRowNum = 1,
            colWidth = 200)
    private String totalProfit;

    @ExcelWriteProperty(titleName = "净利润", rowIndex = 3, colIndex = 14, mergeRowNum = 1,
            colWidth = 200)
    private String netProfit;
    /**
     * 此行数据在set时不需要关注,是固定行内容
     */
    @ExcelWriteProperty(titleName = "银行选填", rowIndex = 3, colIndex = 15, mergeTitleColNum = 2,
            colWidth = 200)
    private String bank;

    @ExcelWriteProperty(titleName = "资本充足率", rowIndex = 4, colIndex = 15,
            colWidth = 200)
    private String capitalAdequacyRat;

    @ExcelWriteProperty(titleName = "拨备覆盖率", rowIndex = 4, colIndex = 16,
            colWidth = 200)
    private String provisionCoverageRat;

    @ExcelWriteProperty(titleName = "不良贷款率", rowIndex = 4, colIndex = 17,
            colWidth = 200)
    private String badLoanRat;
    /**
     * 此行数据在set时不需要关注,是固定行内容
     */
    @ExcelWriteProperty(titleName = "证券公司选填", rowIndex = 3, colIndex = 18, mergeTitleColNum = 1,
            colWidth = 200)
    private String securitiesComp;

    @ExcelWriteProperty(titleName = "融资融券余额", rowIndex = 4, colIndex = 18,
            colWidth = 200)
    private String marginBalance;

    @ExcelWriteProperty(titleName = "净稳定资金率", rowIndex = 4, colIndex = 19,
            colWidth = 200)
    private String netStableCapitalRat;

    /**
     * 此行数据在set时不需要关注,是固定行内容
     */
    @ExcelWriteProperty(titleName = "保险公司选填", rowIndex = 3, colIndex = 20, mergeTitleColNum = 2,
            colWidth = 200)
    private String insuranceComp;

    @ExcelWriteProperty(titleName = "保费收入", rowIndex = 4, colIndex = 20,
            colWidth = 200)
    private String premiumIncome;

    @ExcelWriteProperty(titleName = "核心偿付能力充足率", rowIndex = 4, colIndex = 21,
            colWidth = 200)
    private String coreSolvencyAdequacyRat;

    @ExcelWriteProperty(titleName = "综合偿付能力充足率", rowIndex = 4, colIndex = 22,
            colWidth = 200)
    private String comprehensiveSolvencyAdequacyRat;

    /**
     * 此行数据在set时不需要关注,是固定行内容
     */
    @ExcelWriteProperty(titleName = "信托公司选填", rowIndex = 3, colIndex = 23, mergeTitleColNum = 1,
            colWidth = 200)
    private String trustComp;

    @ExcelWriteProperty(titleName = "存现信托产品规模", rowIndex = 4, colIndex = 23,
            colWidth = 200)
    private String depositTrustProdScale;

    @ExcelWriteProperty(titleName = "风险项目规模", rowIndex = 4, colIndex = 24,
            colWidth = 200)
    private String riskProjectScale;

    /**
     * 此行数据在set时不需要关注,是固定行内容
     */
    @ExcelWriteProperty(titleName = "租赁公司选填", rowIndex = 3, colIndex = 25, mergeTitleColNum = 1,
            colWidth = 200)
    private String leasingComp;

    @ExcelWriteProperty(titleName = "融资租赁资产规模", rowIndex = 4, colIndex = 25,
            colWidth = 200)
    private String leaseAssetsScale;

    @ExcelWriteProperty(titleName = "不良资产率", rowIndex = 4, colIndex = 26,
            colWidth = 200)
    private String badAssetsRat;

    @ExcelWriteProperty(titleName = "金融子企业变化情况及原因", rowIndex = 3, colIndex = 27, mergeRowNum = 1,
            colWidth = 300)
    private String subEntChangesReason;

    @ExcelWriteProperty(titleName = "备注", rowIndex = 3, colIndex = 28, mergeRowNum = 1,
            colWidth = 200)
    private String comments;
}
