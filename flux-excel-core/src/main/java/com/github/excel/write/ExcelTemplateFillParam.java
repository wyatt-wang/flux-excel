package com.github.excel.write;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;
import lombok.experimental.Accessors;

import java.util.List;

/**
 * 模板填充参数
 */
@Data
@Accessors(chain = true)
@AllArgsConstructor
@NoArgsConstructor
public class ExcelTemplateFillParam {
	private String name;
	private Object value;
	private List<?> listValue;
}
