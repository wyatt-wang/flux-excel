package com.github.excel.param;

import lombok.Data;
import lombok.experimental.Accessors;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

/**
 * Excel 富文本单元格值。
 */
@Data
@Accessors(chain = true)
public class ExcelRichTextValue {
	private String text;
	private List<ExcelRichTextModel> richTextModels = new ArrayList<>();

	public static ExcelRichTextValue of(String text, ExcelRichTextModel... richTextModels) {
		ExcelRichTextValue value = new ExcelRichTextValue().setText(text);
		if (richTextModels != null) {
			value.setRichTextModels(new ArrayList<>(Arrays.asList(richTextModels)));
		}
		return value;
	}
}
