package com.github.excel.util;

import com.github.excel.constant.ExcelConstant;
import com.github.excel.exception.ExcelCompressException;
import com.google.common.base.Throwables;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.compress.archivers.zip.*;
import org.apache.commons.compress.parallel.InputStreamSupplier;

import java.io.*;
import java.util.List;
import java.util.Objects;
import java.util.concurrent.ExecutionException;

/**
 * @Author: Vachel Wang
 * @Date: 2026/4/24
 * @Description: 压缩工具类
 */
@Slf4j
public class ZipCompressUtil {
	ParallelScatterZipCreator scatterZipCreator = new ParallelScatterZipCreator();
	ScatterZipOutputStream dirs;
	private static final String TEMP_FILE_PREFIX = "scatter-dirs";
	private static final String TEMP_FILE_SUFFIX = "tmp";
	private static final String ZIP_SUFFIX = ".zip";
	private static final String NULL_STR = "";
	private static final String FILE_SEPARATOR = "/";

	public ZipCompressUtil() {
		try {
			dirs = ScatterZipOutputStream.fileBased(File.createTempFile(TEMP_FILE_PREFIX, TEMP_FILE_SUFFIX));
		} catch (IOException e) {
			throw new ExcelCompressException(e.getMessage());
		}
	}

	public void compressFile(List<File> fileList, String outputPath, String prefixDir) {
		if (StringUtil.isEmpty(outputPath)) {
			throw new ExcelCompressException("outputPath can't be null");
		}
		if (!outputPath.endsWith(ZIP_SUFFIX)) {
			throw new ExcelCompressException("outputPath suffix must to be .zip");
		}
		if (Objects.isNull(prefixDir)) {
			prefixDir = NULL_STR;
		} else {
			if (!prefixDir.endsWith(FILE_SEPARATOR)) {
				prefixDir += FILE_SEPARATOR;
			}
		}
		ZipArchiveOutputStream zipArchiveOutputStream;
		try {
			zipArchiveOutputStream = new ZipArchiveOutputStream(new File(outputPath));
		} catch (IOException e) {
			throw new ExcelCompressException(e.getMessage());
		}
		for (File file : fileList) {
			ZipArchiveEntry archiveEntry = new ZipArchiveEntry(file, prefixDir + file.getName());
			archiveEntry.setMethod(ZipMethod.DEFLATED.getCode());

			InputStreamSupplier supplier = () -> {
				try {
					if (!file.isDirectory()) {
						return new FileInputStream(file);
					} else {
						return new ByteArrayInputStream(new byte[ExcelConstant.ZERO_SHORT]);
					}
				} catch (FileNotFoundException e) {
					throw new ExcelCompressException(e.getMessage());
				}
			};
			addEntry(archiveEntry, supplier);

		}
		writeTo(zipArchiveOutputStream);
	}

	private void addEntry(ZipArchiveEntry zipArchiveEntry, InputStreamSupplier streamSupplier) {
		if (zipArchiveEntry.isDirectory() && !zipArchiveEntry.isUnixSymlink()) {
			try {
				dirs.addArchiveEntry(ZipArchiveEntryRequest.createZipArchiveEntryRequest(zipArchiveEntry, streamSupplier));
			} catch (IOException e) {
				throw new ExcelCompressException(e.getMessage());
			}
		} else {
			scatterZipCreator.addArchiveEntry(zipArchiveEntry, streamSupplier);
		}
	}

	private void writeTo(ZipArchiveOutputStream zipArchiveOutputStream) {
		try {
			dirs.writeTo(zipArchiveOutputStream);
			dirs.close();
			scatterZipCreator.writeTo(zipArchiveOutputStream);
		} catch (IOException e) {
			log.error(Throwables.getStackTraceAsString(e));
			throw new ExcelCompressException(e.getMessage());
		} catch (InterruptedException e) {
			throw new ExcelCompressException(e.getMessage());
		} catch (ExecutionException e) {
			log.error(Throwables.getStackTraceAsString(e));
			throw new ExcelCompressException(e.getMessage());
		}
	}
}
