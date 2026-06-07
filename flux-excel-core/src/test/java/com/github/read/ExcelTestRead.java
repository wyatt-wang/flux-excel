package com.github.read;

import com.github.BaseExcelTest;
import com.github.excel.param.ExcelReaderListParam;
import com.github.excel.param.ExcelReaderStreamParam;
import com.github.excel.read.handler.reader.ExcelReader;
import com.github.excel.read.handler.reader.ExcelReaderFactory;
import com.github.model.UserExcelDtoImportBean;
import com.github.model.UserExcelDtoImportBean1;
import com.github.model.UserExcelDtoImportBean2;
import com.github.model.UserExcelDtoImportBean3;
import com.github.read.handler.ExcelStandardRowHandler;
import com.google.common.base.Stopwatch;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.extractor.ExcelExtractor;
import org.apache.poi.xssf.extractor.XSSFExcelExtractor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;

import java.io.FileInputStream;
import java.io.InputStream;
import java.util.List;
import java.util.concurrent.TimeUnit;


/**
 * @Author: Vachel Wang
 * @Date: 2026/4/24
 * @Description: 简单导出
 */
@Slf4j
public class ExcelTestRead extends BaseExcelTest {

	@Test
    public void testExtra() throws Exception {
		InputStream is = new FileInputStream("/Users/vico/dev-soft/sample-import.xlsx");
		XSSFWorkbook wb = new XSSFWorkbook(is);
		ExcelExtractor extractor = new XSSFExcelExtractor(wb);
		extractor.setFormulasNotResults(false);
		extractor.setIncludeSheetNames(false);
		String text = extractor.getText();
		wb.close();
		log.info(text);

	}
	@Test
	public void testSample() throws Exception {

		InputStream is = new FileInputStream("/Users/vico/dev-soft/sample-import.xlsx");
		ExcelReaderStreamParam streamParam = ExcelReaderStreamParam.builder().stream(is).build();
		ExcelReader reader = ExcelReaderFactory.createUserReader(streamParam);
		ExcelReaderListParam<UserExcelDtoImportBean> listParam = ExcelReaderListParam.<UserExcelDtoImportBean>builder().modelCla(UserExcelDtoImportBean.class).rowHandler(new ExcelStandardRowHandler()).build();
		reader.addList(listParam);
//		reader.addModel(UserExcelDtoImportBean2.class,1,(model -> {log.info("print model:"+model.toString());}));
//		reader.addModel(UserExcelDtoImportBean3.class,1);
		reader.readCustom((workbook -> {
			String val = workbook.getSheetAt(1).getRow(1).getCell(1).getStringCellValue();
			log.info("自定义读取："+ val);
		}));

//		reader.addModelList(UserExcelDtoImportList.class, 0);
		/*reader.addModelList(UserExcelDtoImportList1.class,0, new ExcelReadBatchProcess<UserExcelDtoImportList1>() {
			@Override
			public int getBatchSize() {
				return 3;
			}

			@Override
			public void process(List<UserExcelDtoImportList1> dataList) {
//				throw new ExcelReadException("heh");
				log.info("UserExcelDtoImportList1 ==============");
				dataList.forEach(e-> System.out.println(e));
				log.info("==============");
			}
		});*/

		Stopwatch stopwatch = Stopwatch.createStarted();
		reader.parse();
		List<UserExcelDtoImportBean> modelList = reader.getList(UserExcelDtoImportBean.class);
		List<UserExcelDtoImportBean1> modelList1 = reader.getList(UserExcelDtoImportBean1.class);

		UserExcelDtoImportBean2 model = reader.getModel(UserExcelDtoImportBean2.class);
		UserExcelDtoImportBean3 model3 =  reader.getModel(UserExcelDtoImportBean3.class);
//		reader.parse();
		long elapsed = stopwatch.elapsed(TimeUnit.SECONDS);
		log.info(elapsed + "时间");
		/*log.info("parse exists error: "+errorMsgModel.getExistsError());
		errorMsgModel.getErrorMsgInfoList().forEach(e ->{
			log.error(e.getErrorMsg());
		});*/
		/*UserExcelDtoImportBean model = reader.getModel(UserExcelDtoImportBean.class);
		System.out.println(model);
		UserExcelDtoImportBean1 model1 = reader.getModel(UserExcelDtoImportBean1.class);
		System.out.println(model1);
		UserExcelDtoImportBean2 model2 = reader.getModel(UserExcelDtoImportBean2.class);
		System.out.println(model2);
		List<UserExcelDtoImportList> modelList= reader.getModelList(UserExcelDtoImportList.class);
		List<UserExcelDtoImportList1> modelList1= reader.getModelList(UserExcelDtoImportList1.class);
		modelList.forEach(e -> System.out.println(e));*/
//		modelList1.forEach(e -> System.out.println(e));
		/*int i=1;
		for(ReadPictureModel logo:modelList1.get(4).getLogo()) {
			OutputStream out = new FileOutputStream("/Users/vico/Downloads/test-logo"+i+ logo.getSuffix());
			out.write(logo.getBytes());
			out.close();
			++ i ;
		}*/
	}

	/*@Test
	public void testReadFunction() throws Exception {
		InputStream is = new FileInputStream("/Users/vico/dev-soft/sample-import.xlsx");
		List<UserExcelDtoImportBean> list = ExcelReaderFactory.createUserReader("sample-import.xlsx", is).addList(UserExcelDtoImportBean.class, 0).parseAndGetList(UserExcelDtoImportBean.class);
		log.info("时间");
	}*/


	/*@Test
	public void testReadList() throws Exception {

		InputStream is = new FileInputStream("/Users/vico/Downloads/import-test.xlsx");
		ExcelReader reader = ExcelReaderFactory.createUserReader("import-test.xlsx",is);
		reader.setReadPicture(true);
		reader.addList(UserExcelDtoImportList.class,0);
		reader.addList(UserExcelDtoImportList1.class,0);
//		reader.addModel(UserExcelDtoImportBean1.class,0);
//		reader.addModel(UserExcelDtoImportBean2.class,1);
//		reader.addModelList(UserExcelDtoImportList1.class,1);
		Stopwatch stopwatch = Stopwatch.createStarted();
		reader.parse();
		long elapsed = stopwatch.elapsed(TimeUnit.SECONDS);
		log.info(elapsed + "时间");
		*//*log.info("parse exists error: "+errorMsgModel.getExistsError());
		errorMsgModel.getErrorMsgInfoList().forEach(e ->{
			log.error(e.getErrorMsg());
		});*//*
		List<UserExcelDtoImportList> modelList= reader.getList(UserExcelDtoImportList.class);
		List<UserExcelDtoImportList1> modelList1= reader.getList(UserExcelDtoImportList1.class);
		modelList.forEach(e -> System.out.println(e));
		System.out.println("==============");
		modelList1.forEach(e -> System.out.println(e));
	}
	@Test
	public void testReadProject() throws Exception {

		InputStream is = new FileInputStream("/Users/vico/Documents/flux-excel/导入模板/project-bids.xlsx");
		ExcelReader reader = ExcelReaderFactory.createUserReader("project-bids.xlsx",is,"project-bids.xlsx");
		reader.addList(ProjectBidsBean.class,1);
		Stopwatch stopwatch = Stopwatch.createStarted();
		reader.parse();
		long elapsed = stopwatch.elapsed(TimeUnit.SECONDS);
		log.info(elapsed + "时间");

		List<ProjectBidsBean> modelList = reader.getList(ProjectBidsBean.class);
		for (ProjectBidsBean bidsBean : modelList) {
			System.out.println(bidsBean);
		}
	}

	@Test
	public void testRead() throws Exception {
		InputStream is = new FileInputStream("/Users/vico/Downloads/sample-1.xlsx");
		ExcelReader reader = new ExcelUserReader("sample-1.xlsx",is);
		reader.addModel(UserExcelDtoImportBean.class,0);
		reader.addList(UserExcelDtoImportList.class,0);
		Stopwatch stopwatch = Stopwatch.createStarted();
		reader.parse();
		long elapsed = stopwatch.elapsed(TimeUnit.SECONDS);
		log.info(elapsed + "");
		UserExcelDtoImportBean model = reader.getModel(UserExcelDtoImportBean.class);
		System.out.println(model);
	}

	@Test
	public void testExport() throws Exception {
		byte[] bytes = ExcelBootLoader.getExcelImportTemplateFileCacheMapValue("project-bids.xlsx");
		if (Objects.nonNull(bytes) && bytes.length > ExcelConstant.ZERO_SHORT) {
			OutputStream outputStream = new FileOutputStream(new File("/Users/vico/Downloads/test11111.xlsx"));
			outputStream.write(bytes);
			outputStream.flush();
			outputStream.close();
		}

	}
	@Test
	public void testNumberRegex() throws Exception {
		boolean flag = ".2341".matches(ExcelConstant.NUMBER_PATTERN);
		System.out.println(flag);
		//43389.64541107639
		//1540199621691
		System.out.println(System.currentTimeMillis());
	}*/
}
