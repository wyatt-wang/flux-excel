package com.github.excel.param;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;
import lombok.experimental.Accessors;
import lombok.experimental.SuperBuilder;
import org.apache.poi.ss.usermodel.RichTextString;

/**
 * @Author: Vachel Wang
 * @Date: 2026/4/24
 * @Description: 批注model
 */
@Data
@Accessors(chain = true)
@SuperBuilder
@AllArgsConstructor
@NoArgsConstructor
public class ExcelWriterCommentParam{
	/**
	 * 当前单元格的原始值
	 */
	private Object value;
	/**
	 * 兼容旧版字符串批注
	 */
	private String commentText;
	/**
	 * 批注字体名称
	 */
	private String commentFontName;
	/**
	 * 批注
	 */
	private RichTextString comment;
	/**
	 * 批注作者
	 */
	private String author;
}
