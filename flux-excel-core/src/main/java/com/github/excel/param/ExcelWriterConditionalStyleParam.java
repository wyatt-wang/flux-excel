package com.github.excel.param;

import com.github.excel.enums.ExcelWriterConditionOperatorEnum;
import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;
import lombok.experimental.Accessors;
import lombok.experimental.SuperBuilder;
import org.apache.poi.ss.usermodel.IndexedColors;

/**
 * Excel 条件样式参数
 */
@Data
@Accessors(chain = true)
@SuperBuilder
@AllArgsConstructor
@NoArgsConstructor
public class ExcelWriterConditionalStyleParam extends ExcelWriterDataParam {
	/**
	 * 结束行号
	 */
	private Integer endRowIndex;
	/**
	 * 结束列号
	 */
	private Integer endColIndex;
	/**
	 * 自定义公式。配置后优先使用公式创建条件样式。
	 */
	private String formula;
	/**
	 * 比较操作符。
	 */
	private ExcelWriterConditionOperatorEnum operator = ExcelWriterConditionOperatorEnum.NO_COMPARISON;
	/**
	 * 条件值1。
	 */
	private String start;
	/**
	 * 条件值2，BETWEEN/NOT_BETWEEN 时使用。
	 */
	private String end;
	/**
	 * 前景填充色。
	 */
	private Short fillForegroundColor;
	/**
	 * 字体颜色。
	 */
	private Short fontColor;
	/**
	 * 是否加粗。
	 */
	private Boolean bold;
	/**
	 * 是否斜体。
	 */
	private Boolean italic;

	public static ExcelWriterConditionalStyleParam greaterThan(String value) {
		return new ExcelWriterConditionalStyleParam()
				.setOperator(ExcelWriterConditionOperatorEnum.GT)
				.setStart(value);
	}

	public static ExcelWriterConditionalStyleParam lessThan(String value) {
		return new ExcelWriterConditionalStyleParam()
				.setOperator(ExcelWriterConditionOperatorEnum.LT)
				.setStart(value);
	}

	public static ExcelWriterConditionalStyleParam formula(String formula) {
		return new ExcelWriterConditionalStyleParam().setFormula(formula);
	}

	public ExcelWriterConditionalStyleParam fill(IndexedColors color) {
		if (color != null) {
			this.fillForegroundColor = color.getIndex();
		}
		return this;
	}

	public ExcelWriterConditionalStyleParam font(IndexedColors color) {
		if (color != null) {
			this.fontColor = color.getIndex();
		}
		return this;
	}
}
