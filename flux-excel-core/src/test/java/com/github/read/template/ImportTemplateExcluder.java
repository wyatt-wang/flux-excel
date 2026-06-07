package com.github.read.template;

import com.github.excel.read.facade.AbstractReaderTemplateExclude;

/**
 * @Author: Vachel Wang
 * @Date: 2026/4/24
 * @Description: 简单导出
 */
public class ImportTemplateExcluder extends AbstractReaderTemplateExclude {
	@Override
	public void addTemplateExclude() {
		this.addExclude(new SheetExclude("说明", "project-bids.xlsx"))
				.addExclude(new RowExclude(1, "招标规划", "project-bids.xlsx"))
				.addExclude(new ColumnExclude(1,2,"招标规划","project-bids.xlsx"));
	}
}
