package com.github.excel.write.impl;

import com.github.excel.boot.ExcelBootLoader;
import com.github.excel.constant.ExcelConstant;
import com.github.excel.exception.ExcelWriterException;
import com.github.excel.helper.WorkbookHelper;
import com.github.excel.model.ExcelBaseModel;
import com.github.excel.model.ExcelCacheModel;
import com.github.excel.model.ExcelCacheFieldModel;
import com.github.excel.util.ExcelUtil;
import com.github.excel.util.StringUtil;
import com.github.excel.util.ZipCompressUtil;
import com.github.excel.write.*;
import com.github.excel.write.style.AbstractExcelStyle;
import com.github.excel.write.style.ExcelBasicStyle;
import com.google.common.base.Throwables;
import com.google.common.collect.Lists;
import lombok.Data;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.collections4.CollectionUtils;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import jakarta.servlet.ServletOutputStream;
import jakarta.servlet.http.HttpServletRequest;
import jakarta.servlet.http.HttpServletResponse;
import java.io.*;
import java.util.*;
import java.util.concurrent.*;
import java.util.concurrent.atomic.AtomicLong;

/**
 * @Author: Vachel Wang
 * @Date: 2026/4/24
 * @Description: 超大list异步导出多个文件，并压缩
 */
@Slf4j
public class ExcelLargeListBatchWriterImpl extends BaseExcelWriter implements ExcelLargeListBatchWriter {


	private final ExecutorService executorService;

	private String outputDirPath;

	private final List<LargeListAsyncParam> asyncParamList = Lists.newArrayList();

	private final File dirFile;

	private static final String ZIP_SUFFIX = ".zip";

	private final int MAX_POOL_SIZE = 20 ;

	private static final String LARGE_LIST_BATCH_POOL = "ExcelLargeListBatchPool-";

	private final int rowAccessWindowSize;
	private final boolean compressTempFiles;
	private final boolean useSharedStringsTable;

	public ExcelLargeListBatchWriterImpl(String outputDirPath) {
		this(outputDirPath, null);
	}

	public ExcelLargeListBatchWriterImpl(String outputDirPath ,Integer maxPoolSize) {
		this(outputDirPath, maxPoolSize, ExcelConstant.INT_10000, false, false);
	}

	public ExcelLargeListBatchWriterImpl(String outputDirPath, Integer maxPoolSize, int rowAccessWindowSize,
										 boolean compressTempFiles, boolean useSharedStringsTable) {
		if (StringUtil.isEmpty(outputDirPath)) {
			throw new ExcelWriterException("outputDirPath can't be null ");
		}
		if (Objects.isNull(maxPoolSize)) {
			maxPoolSize = MAX_POOL_SIZE;
		}
		dirFile = new File(outputDirPath);
		if (!dirFile.exists()) {
			dirFile.mkdirs();
		}
		this.outputDirPath = outputDirPath;
		this.rowAccessWindowSize = rowAccessWindowSize;
		this.compressTempFiles = compressTempFiles;
		this.useSharedStringsTable = useSharedStringsTable;
		this.addStyle(ExcelBasicStyle.class);
		executorService = new ThreadPoolExecutor(ExcelConstant.ONE_INT, maxPoolSize, ExcelConstant.ONE_INT, TimeUnit.MINUTES, new SynchronousQueue<>(), new ThreadFactory() {
			int i = ExcelConstant.ZERO_SHORT;

			@Override
			public Thread newThread(Runnable r) {
				i++;
				return new Thread(r,LARGE_LIST_BATCH_POOL + i);
			}
		});
	}

	@Override
	public <T extends ExcelBaseModel> void process(List<T> modelList, String fileName, String sheetName,String[] excludeFields,Class<? extends ExcelBaseModel> modelCla) {
		if (StringUtil.isEmpty(fileName)) {
			throw new ExcelWriterException("fileName can't be null");
		}
		SXSSFWorkbook workbook = WorkbookHelper.createStreamingXlsxWorkBook(rowAccessWindowSize, compressTempFiles,
				useSharedStringsTable);
		String excelPath = outputDirPath + ExcelConstant.FILE_SEPARATOR + fileName + ExcelConstant.XLSX_STR;
		LargeListAsyncParam param = new LargeListAsyncParam(sheetName, fileName, modelList, workbook, outputDirPath, excelPath,excludeFields,modelCla);
		LargeListAsyncWriterTask writerTask = new LargeListAsyncWriterTask(param);
		Future<Boolean> future = executorService.submit(writerTask);
		param.setFuture(future);
		asyncParamList.add(param);

	}

	@Override
	public <T extends ExcelBaseModel> void process(List<T> modelList, String fileName, String sheetName,Class<? extends ExcelBaseModel> modelCla) {
		process(modelList,fileName,sheetName,null,modelCla);
	}

	@Override
	public void export(String zipFileName) {
		if (StringUtil.isEmpty(zipFileName)) {
			throw new ExcelWriterException("zipFileName can't be null");
		}
		zipFileName+=ZIP_SUFFIX;
		try {
			List<File> fileList = Lists.newArrayList();
			for (int i = ExcelConstant.ZERO_SHORT; i < asyncParamList.size(); i++) {
				LargeListAsyncParam param = asyncParamList.get(i);
				Boolean result = param.getFuture().get();
				if (Objects.isNull(result) || !result) {
					throw new ExcelWriterException("export excel failed. cause:fill batch no " + i + " error");
				}
				File file = new File(param.getExcelPath());
				fileList.add(file);
			}

			ZipCompressUtil compressUtil = new ZipCompressUtil();
			compressUtil.compressFile(fileList, dirFile + ExcelConstant.FILE_SEPARATOR + zipFileName, null);
			for (File file : fileList) {
				file.delete();
			}
		} catch (InterruptedException e) {
			Thread.currentThread().interrupt();
			log.error("export failed cause:{}", Throwables.getStackTraceAsString(e));
			throw new ExcelWriterException(e.getMessage());
		} catch (ExecutionException e) {
			log.error("export failed cause:{}", Throwables.getStackTraceAsString(e));
			throw new ExcelWriterException(e.getMessage());
		}finally {
			executorService.shutdown();
			clearData();
		}
	}

	@Override
	public void export(HttpServletRequest request, HttpServletResponse response, String zipFileName) {
		try {
			export(zipFileName);
			ExcelUtil.setResponseHeader(request, response, zipFileName, null);
			ServletOutputStream outputStream = response.getOutputStream();
			zipFileName+=ZIP_SUFFIX;
			File zipFile = new File(dirFile + ExcelConstant.FILE_SEPARATOR + zipFileName);
			try (InputStream inputStream = new FileInputStream(zipFile)) {
				byte[] buffer = new byte[8192];
				int length;
				while ((length = inputStream.read(buffer)) != -1) {
					outputStream.write(buffer, 0, length);
				}
				outputStream.flush();
			}
		} catch (UnsupportedEncodingException e) {
			log.error("export failed cause:{}", Throwables.getStackTraceAsString(e));
			throw new ExcelWriterException(e.getMessage());
		} catch (IOException e) {
			log.error("export failed cause:{}", Throwables.getStackTraceAsString(e));
			throw new ExcelWriterException(e.getMessage());
		}finally {
			clearData();
		}
	}

	@Override
	public void addStyle(Class<? extends AbstractExcelStyle> styleClass) {
		this.styleList.add(styleClass);
	}

	/**
	 * @Description: 异步导出任务
	 * @Author: Vachel Wang
	 * @Date: 2026/4/24
	 * @Email:
	 */
	private class LargeListAsyncWriterTask implements Callable<Boolean> {
		private LargeListAsyncParam param;

		public LargeListAsyncWriterTask(LargeListAsyncParam param) {
			this.param = param;
		}

		@Override
		public Boolean call() throws Exception {
			SXSSFWorkbook workbook = param.getWorkbook();
			CreationHelper creationHelper = workbook.getCreationHelper();
			try {
				// 初始化样式
				initStyle(workbook);
				if (CollectionUtils.isEmpty(param.getModelList())) {
					// 判断是否有数据，没有数据给出默认提示
					addNoResultData(workbook, creationHelper);
				} else {
					Class<? extends ExcelBaseModel> modelClass = param.getModelCla();
					ExcelCacheModel cacheModel = ExcelBootLoader.getExcelCacheMapValue(modelClass);

					int rowIndex = ExcelConstant.ZERO_SHORT;
					String sheetName = param.getSheetName();
					String[] excludeFields = param.getExcludeFields();
					List<ExcelCacheFieldModel> cacheFieldModelList = filterListFieldModels(cacheModel, excludeFields);
					// 填充标题
					SXSSFSheet sheet = workbook.createSheet(sheetName);
					writeFlatListTitleRow(sheet, creationHelper, cacheFieldModelList, rowIndex);
					AtomicLong incrementSeq = new AtomicLong(ExcelConstant.ZERO_SHORT);
					++rowIndex;
					// 填充内容
					int i = ExcelConstant.ZERO_SHORT ;
					for (ExcelBaseModel model : param.getModelList()) {
						i++;
						writeFlatListContentRow(workbook, sheet, creationHelper, cacheFieldModelList, model, incrementSeq, rowIndex, i);
						rowIndex++;
					}
				}
				// 保存内容
				try (OutputStream outputStream = new FileOutputStream(new File(param.getExcelPath()))) {
					workbook.write(outputStream);
				}
			} catch (Exception ex) {
				log.error("export failed thread:{} ,cause:{}", Thread.currentThread().getName(), Throwables.getStackTraceAsString(ex));
				return Boolean.FALSE;
			}finally {
				workbook.close();
				workbook.dispose();
				styleLocal.remove();
				fontLocal.remove();
			}
			return Boolean.TRUE;
		}
	}


	/**
	 * @Description: 异步导出参数
	 * @Author: Vachel Wang
	 * @Date: 2026/4/24
	 * @Email:
	 */
	@Data
	private class LargeListAsyncParam {

		private String sheetName;

		private String fileName;

		private List<? extends ExcelBaseModel> modelList;

		private volatile SXSSFWorkbook workbook;

		private String outputDirPath;

		private String excelPath;

		private Future<Boolean> future;

		private String[] excludeFields;

		private Class<? extends ExcelBaseModel> modelCla ;

		public LargeListAsyncParam(String sheetName, String fileName, List<? extends ExcelBaseModel> modelList, SXSSFWorkbook workbook, String outputDirPath, String excelPath, String[] excludeFields,Class<? extends ExcelBaseModel> modelCla) {
			this.sheetName = sheetName;
			this.fileName = fileName;
			this.modelList = modelList;
			this.workbook = workbook;
			this.outputDirPath = outputDirPath;
			this.excelPath = excelPath;
			this.excludeFields = excludeFields;
			this.modelCla = modelCla;
		}
	}

}
