package com.github.excel.param;

import com.github.excel.model.ExcelBaseModel;
import com.github.excel.validator.ExcelValidationGroup;
import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;
import lombok.experimental.Accessors;
import lombok.experimental.SuperBuilder;
import org.apache.poi.ss.formula.functions.T;
import org.hibernate.validator.constraints.Range;

import jakarta.validation.constraints.NotEmpty;
import jakarta.validation.constraints.NotNull;

/**
 * @Author: Vachel Wang
 * @Date: 2026/4/24
 * @Description: 导出到单个单元格model
 */
@Data
@Accessors(chain = true)
@SuperBuilder
@AllArgsConstructor
@NoArgsConstructor
public class ExcelWriterMergeParam extends ExcelWriterCellParam {
	/**
	 * 结束行号
	 */
	@NotNull(message = "结束行号不能为空")
	@Range(min = 0, max = 1048576 , message = "结束行号不能小于0且在xlsx下不能大于1048576" ,groups = {ExcelValidationGroup.XssfWorkbook.class})
	@Range(min = 0, max = 1048576 , message = "结束行号不能小于0且在xlsx下不能大于1048576" ,groups = {ExcelValidationGroup.SxssfWorkbook.class})
	@Range(min = 0, max = 65536 , message = "结束行号不能小于0且在xls下不能大于65536" ,groups = {ExcelValidationGroup.HssfWorkbook.class})
	private Integer endRowIndex;
	/**
	 * 结束列号
	 */
	@NotNull(message = "结束列号不能为空")
	@Range(min = 0, max = 16384, message = "结束列号不能小于0且在xlsx下不能大于16384", groups = {ExcelValidationGroup.XssfWorkbook.class})
	@Range(min = 0, max = 16384, message = "结束列号不能小于0且在xlsx下不能大于16384", groups = {ExcelValidationGroup.SxssfWorkbook.class})
	@Range(min = 0, max = 255, message = "结束列号不能小于0且在xls下不能大于255", groups = {ExcelValidationGroup.HssfWorkbook.class})
	private Integer endColIndex;

	public ExcelWriterMergeParam setFirstRow(Integer firstRow) {
		setRowIndex(firstRow);
		return this;
	}

	public ExcelWriterMergeParam setFirstColumn(Integer firstColumn) {
		setColIndex(firstColumn);
		return this;
	}

	public ExcelWriterMergeParam setLastRow(Integer lastRow) {
		setEndRowIndex(lastRow);
		return this;
	}

	public ExcelWriterMergeParam setLastColumn(Integer lastColumn) {
		setEndColIndex(lastColumn);
		return this;
	}
}
