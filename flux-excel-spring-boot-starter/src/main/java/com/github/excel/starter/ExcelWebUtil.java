package com.github.excel.starter;

import com.github.excel.boot.ExcelBootLoader;
import com.github.excel.constant.ExcelConstant;
import com.github.excel.util.StringUtil;
import com.google.common.base.Throwables;
import jakarta.servlet.ServletOutputStream;
import jakarta.servlet.http.HttpServletRequest;
import jakarta.servlet.http.HttpServletResponse;
import lombok.extern.slf4j.Slf4j;

import java.io.IOException;
import java.io.PrintWriter;
import java.io.UnsupportedEncodingException;
import java.net.URLEncoder;
import java.util.Objects;

@Slf4j
public final class ExcelWebUtil {

	private ExcelWebUtil() {
	}

	public static void setResponseHeader(HttpServletRequest request, HttpServletResponse response,
										 String fileName, String suffix) throws UnsupportedEncodingException {
		if (StringUtil.notEmpty(suffix) && !fileName.endsWith(suffix)) {
			fileName = fileName + suffix;
		}
		String userAgent = request.getHeader("USER-AGENT");
		String finalFileName;
		if (userAgent != null && (userAgent.contains("MSIE") || userAgent.contains("Trident"))) {
			finalFileName = URLEncoder.encode(fileName, "UTF-8");
		} else {
			finalFileName = new String(fileName.getBytes("UTF-8"), "ISO-8859-1");
		}
		response.setCharacterEncoding("UTF-8");
		response.setHeader("Content-Disposition", "attachment; filename=\"" + finalFileName + "\"");
		response.setContentType("application/vnd.ms-excel;charset=utf-8");
	}

	public static void setResponseHeader(HttpServletResponse response, String fileName) {
		if (StringUtil.isEmpty(fileName) || Objects.isNull(response)) {
			return;
		}
		if (fileName.endsWith(ExcelConstant.XLSX_STR)) {
			response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
		} else if (fileName.endsWith(ExcelConstant.XLS_STR)) {
			response.setContentType("application/vnd.ms-excel");
		}
		try {
			fileName = URLEncoder.encode(fileName, "UTF8");
		} catch (UnsupportedEncodingException e) {
			log.error("Download import template error , cause:{}", Throwables.getStackTraceAsString(e));
		}
		response.setHeader("Content-Disposition", "attachment; filename=\"" + fileName + "\"");
	}

	public static void downloadImportTemplate(HttpServletResponse response, String templateName, String alias) {
		byte[] bytes = ExcelBootLoader.getExcelImportTemplateFileCacheMapValue(templateName);
		if (Objects.nonNull(bytes) && bytes.length > ExcelConstant.ZERO_SHORT) {
			if (StringUtil.isEmpty(alias)) {
				alias = templateName;
			} else {
				alias += templateName.substring(templateName.indexOf(ExcelConstant.DOT_CHAR));
			}
			setResponseHeader(response, alias);
			try {
				ServletOutputStream outputStream = response.getOutputStream();
				outputStream.write(bytes);
				outputStream.flush();
			} catch (IOException e) {
				log.error("Download import template error , cause:{}", Throwables.getStackTraceAsString(e));
			}
			return;
		}
		try {
			response.setHeader("Content-type", "text/html;charset=UTF-8");
			response.setCharacterEncoding("UTF-8");
			PrintWriter writer = response.getWriter();
			writer.print("<h1>下载模板失败，请确认模板【" + templateName + "】是否添加到导入模板文件目录！</h1>");
			writer.flush();
		} catch (IOException e) {
			log.error("Download import template error , cause:{}", Throwables.getStackTraceAsString(e));
		}
	}
}
