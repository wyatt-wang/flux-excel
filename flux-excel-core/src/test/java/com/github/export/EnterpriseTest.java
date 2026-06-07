package com.github.export;

import com.github.BaseExcelTest;
import com.github.excel.enums.ExcelSuffixEnum;
import com.github.excel.write.ExcelWriter;
import com.github.excel.write.ExcelWriterFactory;
import com.github.model2.Accessory1ModelDTO;
import com.github.model2.Accessory1TitleDTO;
import com.github.model2.ReportTitleDTO;
import lombok.extern.slf4j.Slf4j;
import org.junit.Test;

import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.List;


/**
 * @Author: Vachel Wang
 * @Date: 2026/4/24
 * @Description: excel 测试
 */
@Slf4j
public class EnterpriseTest extends BaseExcelTest {


	@Test
	public void test() throws Exception {
		ExcelWriter excelWriter = ExcelWriterFactory.createUserModelWriter();
		List<Accessory1ModelDTO> modelDTOS = new ArrayList<>();
		Accessory1ModelDTO dto = new Accessory1ModelDTO();
		dto.setEquityLevel("1");
		modelDTOS.add(dto);

		excelWriter.addModel(new Accessory1TitleDTO(), "sheet", false, Accessory1TitleDTO.class);
		excelWriter.addModelList(5,0,modelDTOS, "sheet", false, Accessory1ModelDTO.class);
		excelWriter.process(new FileOutputStream("/Users/vico/Downloads/enterprise.xlsx"), "enterprise.xlsx", ExcelSuffixEnum.XLSX);

	}

	@Test
	public void testReportTitle() throws Exception {
		ExcelWriter excelWriter = ExcelWriterFactory.createUserModelWriter();
		excelWriter.addModel(new ReportTitleDTO(), "sheet", false, ReportTitleDTO.class);
		excelWriter.process(new FileOutputStream("/Users/vico/Downloads/report-title.xlsx"), "report-title.xlsx", ExcelSuffixEnum.XLSX);

	}
}

