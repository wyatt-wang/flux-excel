package com.github.fluent;

import com.github.BaseExcelTest;
import com.github.excel.Excel;
import com.github.excel.enums.ExcelSuffixEnum;
import com.github.model.UserReadAndExportDto;
import com.github.model2.CompanyDto2;
import org.junit.Assert;
import org.junit.Test;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.List;

public class ExcelFluentApiTest extends BaseExcelTest {

	@Test
	public void writeAndReadListWithFluentApi() throws Exception {
		File file = File.createTempFile("flux-excel-fluent-", ".xlsx");
		List<UserReadAndExportDto> users = new ArrayList<>();
		for (int i = 1; i <= 3; i++) {
			users.add(UserReadAndExportDto.builder()
					.name("张三" + i)
					.sex((byte) 1)
					.age(i)
					.height(170F + i)
					.build());
		}

		try (FileOutputStream outputStream = new FileOutputStream(file)) {
			Excel.write(outputStream)
					.fileName(file.getName())
					.suffix(ExcelSuffixEnum.XLSX)
					.sheet("sheet")
					.list(users, UserReadAndExportDto.class)
					.export();
		}

		List<UserReadAndExportDto> parsed;
		try (FileInputStream inputStream = new FileInputStream(file)) {
			parsed = Excel.read(inputStream)
					.fileName(file.getName())
					.closeInputStream(true)
					.sheet(0)
					.list(UserReadAndExportDto.class)
					.parse()
					.getList(UserReadAndExportDto.class);
		}

		Assert.assertEquals(3, parsed.size());
		Assert.assertEquals("张三1", parsed.get(0).getName());
		Assert.assertEquals(Integer.valueOf(1), parsed.get(0).getAge());
	}

	@Test
	public void largeWriteWithFluentApi() throws Exception {
		File file = File.createTempFile("flux-excel-large-fluent-", ".xlsx");
		List<CompanyDto2> companies = new ArrayList<>();
		for (int i = 1; i <= 2; i++) {
			CompanyDto2 company = new CompanyDto2();
			company.setName("公司" + i);
			company.setAddress("address" + i);
			company.setPersons(10 + i);
			company.setCreateTime(new Date());
			company.setUpdateTime(Calendar.getInstance());
			companies.add(company);
		}

		try (FileOutputStream outputStream = new FileOutputStream(file)) {
			Excel.largeWrite(outputStream)
					.sheet("数据导出")
					.sheetRowMaxCount(100)
					.modelClass(CompanyDto2.class)
					.list(companies)
					.export()
					.close();
		}

		Assert.assertTrue(file.length() > 0);
	}
}
