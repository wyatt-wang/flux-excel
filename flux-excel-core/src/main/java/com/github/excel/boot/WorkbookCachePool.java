package com.github.excel.boot;

import com.github.excel.constant.ExcelConstant;
import com.github.excel.exception.ExcelWriterException;
import com.github.excel.write.style.AbstractExcelStyle;
import com.github.excel.write.style.ExcelBasicStyle;
import com.google.common.base.Throwables;
import lombok.Getter;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Color;
import org.apache.poi.ss.usermodel.Workbook;

import java.lang.reflect.Constructor;
import java.util.HashMap;
import java.util.Map;

/**
 * todo 类重构
 * @Author: Vachel Wang
 * @Date: 2026/4/24
 * @Description: Workbook缓存池
 */
@Slf4j
public class WorkbookCachePool {

	@Getter
	private static ThreadLocal<WorkbookCacheModel> hssfWorkbookThreadLocal;

	@Getter
	private static ThreadLocal<WorkbookCacheModel> xssfWorkbookThreadLocal;

	@Getter
	private static ThreadLocal<WorkbookCacheModel> sxssfWorkbookThreadLocal;

	/**
	 * 工作簿解析器
	 */
	public interface WorkbookResolver {
		Workbook resolve();
	}

	/**
	 * 工作簿样式缓存模型
	 */
	@Getter
	public static class WorkbookCacheModel {
		private final Workbook workbook;
		private final Map<String, CellStyle> styleMap;
		private final Map<String, Font> fontMap;
		private final Map<String, Color> colorMap;

		public WorkbookCacheModel(Workbook workbook, Map<String, CellStyle> styleMap, Map<String, Font> fontMap, Map<String, Color> colorMap) {
			this.workbook = workbook;
			this.styleMap = styleMap;
			this.fontMap = fontMap;
			this.colorMap = colorMap;
		}
	}

	/**
	 * 添加style
	 *
	 * @param workbook workbook
	 * @return
	 */
	public static ThreadLocal<WorkbookCacheModel> addBasicStyle(Workbook workbook) {
		try {
			WorkbookCacheModel workbookCacheModel = createCacheModel(workbook);
			return ThreadLocal.withInitial(() -> {
				return workbookCacheModel;
			});
		} catch (Exception e) {
			log.error(Throwables.getStackTraceAsString(e));
			throw new ExcelWriterException("Init style error");
		}
	}

	public static WorkbookCacheModel createCacheModel(Workbook workbook) {
		try {
			Class<ExcelBasicStyle> basicStyleClass = ExcelBasicStyle.class;
			Map<String, CellStyle> styleMap = new HashMap<>();
			Map<String, Font> fontMap = new HashMap<>();
			Map<String, Color> colorMap = new HashMap<>();
			Constructor<? extends AbstractExcelStyle> constructor = basicStyleClass.getConstructor(basicStyleClass.getConstructors()[ExcelConstant.ZERO_SHORT].getParameterTypes());
			AbstractExcelStyle excelStyle = constructor.newInstance(workbook, styleMap, fontMap, colorMap);
			excelStyle.addNewFont();
			excelStyle.addNewStyle();
			excelStyle.addNewColor();
			WorkbookCacheModel workbookCacheModel = new WorkbookCacheModel(workbook, styleMap, fontMap, colorMap);
			return workbookCacheModel;
		} catch (Exception e) {
			log.error(Throwables.getStackTraceAsString(e));
			throw new ExcelWriterException("Init style error");
		}
	}

	/**
	 * 添加style
	 *
	 * @return
	 */
	public static ThreadLocal<WorkbookCacheModel> addBasicStyle(WorkbookResolver workbookResolver) {

		Class<ExcelBasicStyle> basicStyleClass = ExcelBasicStyle.class;

		return ThreadLocal.withInitial(() -> {
			Workbook workbook = workbookResolver.resolve();
			try {
				Map<String, CellStyle> styleMap = new HashMap<>();
				Map<String, Font> fontMap = new HashMap<>();
				Map<String, Color> colorMap = new HashMap<>();
				Constructor<? extends AbstractExcelStyle> constructor = basicStyleClass.getConstructor(basicStyleClass.getConstructors()[ExcelConstant.ZERO_SHORT].getParameterTypes());
				AbstractExcelStyle excelStyle = constructor.newInstance(workbook, styleMap, fontMap, colorMap);
				excelStyle.addNewFont();
				excelStyle.addNewStyle();
				excelStyle.addNewColor();
				return new WorkbookCacheModel(workbook, styleMap, fontMap, colorMap);
			} catch (Exception e) {
				log.error(Throwables.getStackTraceAsString(e));
				throw new ExcelWriterException("Init style error");
			}
		});

	}

}
