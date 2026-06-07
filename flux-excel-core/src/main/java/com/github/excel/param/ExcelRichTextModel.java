package com.github.excel.param;

import lombok.Data;
import org.apache.poi.ss.usermodel.Font;

/**
 * @Author: Vachel Wang
 * @Date: 2026/4/24
 * @Description: Excel 富文本
 */
@Data
public class ExcelRichTextModel {
	private Font font ;
	private int startIndex ;
	private int endIndex ;
}
