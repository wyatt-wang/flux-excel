package com.github.excel.model;

import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;

/**
 * @Author: Vachel Wang
 * @Date: 2026/4/24
 * @Description: 模板标题模型
 */
@Data
@Builder
@AllArgsConstructor
public class ExcelTemplateTitleModel {

    private String sheetName;

    private Integer rowIndex;

    private  Integer colIndex;

    private String title;

}
