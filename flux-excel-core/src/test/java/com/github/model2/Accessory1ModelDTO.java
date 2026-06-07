package com.github.model2;

import com.github.excel.annotation.ExcelWrite;
import com.github.excel.annotation.ExcelWriteProperty;
import com.github.excel.enums.ExcelThemeEnum;
import com.github.excel.model.ExcelBaseModel;
import lombok.Data;

/**
 * @ClassName: Accessory1TitleDTO
 * @Description: 附件1DTO
 * @Author: Vachel Wang
 * @Date: 2026/4/24
 **/
@Data
@ExcelWrite(themeName = ExcelThemeEnum.ZEBRA, fillTitle = false, incrementSequenceNo = true, incrementSequenceTitle = "行次")
public class Accessory1ModelDTO extends ExcelBaseModel {

    @ExcelWriteProperty(titleName = "集团标准全称", colWidth = 350)
    private String groupName;

    @ExcelWriteProperty(titleName = "金融子企业名称", colWidth = 350)
    private String enterpriseName;

    @ExcelWriteProperty(titleName = "金融行业类型-国务院国资委", colWidth = 150)
    private String financialIndustry;

    @ExcelWriteProperty(titleName = "金融机构二级分类", colWidth = 150)
    private String financeSecType;

    @ExcelWriteProperty(titleName = "修改原因", colWidth = 200)
    private String modReason;

    @ExcelWriteProperty(titleName = "合计持股比例", colWidth = 250)
    private String shareHoldingRat;

    @ExcelWriteProperty(titleName = "主要股东标准全称及持股比例", colWidth = 250)
    private String largestShareholderNameAndRat;

    @ExcelWriteProperty(titleName = "管理层级", colWidth = 150)
    private String managementLevel;

    @ExcelWriteProperty(titleName = "股权层级", colWidth = 150)
    private String equityLevel;

    @ExcelWriteProperty(titleName = "总资产", colWidth = 150)
    private String totalAssets;

    @ExcelWriteProperty(titleName = "净资产", colWidth = 150)
    private String netAssets;

    @ExcelWriteProperty(titleName = "营业收入", colWidth = 150)
    private String income;

    @ExcelWriteProperty(titleName = "利润总额", colWidth = 150)
    private String totalProfit;

    @ExcelWriteProperty(titleName = "净利润", colWidth = 150)
    private String netProfit;

    @ExcelWriteProperty(titleName = "资本充足率", colWidth = 150)
    private String capitalAdequacyRat;

    @ExcelWriteProperty(titleName = "拨备覆盖率", colWidth = 150)
    private String provisionCoverageRat;

    @ExcelWriteProperty(titleName = "不良贷款率", colWidth = 150)
    private String badLoanRat;

    @ExcelWriteProperty(titleName = "融资融券余额", colWidth = 150)
    private String marginBalance;

    @ExcelWriteProperty(titleName = "净稳定资金率", colWidth = 150)
    private String netStableCapitalRat;

    @ExcelWriteProperty(titleName = "保费收入", colWidth = 150)
    private String premiumIncome;

    @ExcelWriteProperty(titleName = "核心偿付能力充足率", colWidth = 150)
    private String coreSolvencyAdequacyRat;

    @ExcelWriteProperty(titleName = "综合偿付能力充足率", colWidth = 150)
    private String comprehensiveSolvencyAdequacyRat;

    @ExcelWriteProperty(titleName = "存现信托产品规模", colWidth = 150)
    private String depositTrustProdScale;

    @ExcelWriteProperty(titleName = "风险项目规模", colWidth = 150)
    private String riskProjectScale;

    @ExcelWriteProperty(titleName = "融资租赁资产规模", colWidth = 150)
    private String leaseAssetsScale;

    @ExcelWriteProperty(titleName = "不良资产率", colWidth = 150)
    private String badAssetsRat;

    @ExcelWriteProperty(titleName = "金融子企业变化情况及原因", colWidth = 250)
    private String subEntChangesReason;

    @ExcelWriteProperty(titleName = "备注", colWidth = 300)
    private String comments;

    //是否确认
    private Integer confirm;
    //集团统一社会信用代码
    private String groupCode;
    //启用状态:0-停用, 1-启用
    private Integer enableStatus;

}
