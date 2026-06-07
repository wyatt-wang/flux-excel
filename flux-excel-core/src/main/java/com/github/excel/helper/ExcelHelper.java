package com.github.excel.helper;

import com.github.excel.constant.ExcelConstant;
import com.github.excel.constant.ExcelErrorMsgConstant;
import com.github.excel.exception.ExcelReaderException;
import com.github.excel.exception.ExcelWriterException;
import com.github.excel.param.ExcelRichTextModel;
import com.github.excel.model.ExcelReaderPictureModel;
import com.github.excel.util.StringUtil;
import org.apache.poi.common.usermodel.HyperlinkType;
import org.apache.poi.hssf.usermodel.DVConstraint;
import org.apache.poi.hssf.usermodel.HSSFPalette;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.usermodel.HeaderFooter;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.xssf.usermodel.XSSFDataValidation;

import java.util.*;
import java.util.regex.Pattern;

/**
 * @Author: Vachel Wang
 * @Date: 2026/4/24
 * @Description: excel 帮助类
 */
public class ExcelHelper {

	public static final DataFormatter DATA_FORMATTER = new DataFormatter();

	public static final Pattern NUMBER_PATTERN = Pattern.compile(ExcelConstant.NUMBER_PATTERN_STR);


	/**
	 * 创建富文本
	 *
	 * @param creationHelper    创建
	 * @param text              文本
	 * @param richTextModelList 富文本属性List
	 * @return RichTextString
	 */
	public static RichTextString createRichText(CreationHelper creationHelper, String text, ExcelRichTextModel... richTextModelList) {
		RichTextString richTextString = creationHelper.createRichTextString(text);
		for (ExcelRichTextModel textModel : richTextModelList) {
			if (null == textModel) {
				continue;
			}
			richTextString.applyFont(textModel.getStartIndex(), textModel.getEndIndex(), textModel.getFont());
		}
		return richTextString;
	}

	/**
	 * 获取超链接
	 *
	 * @param createHelper helper
	 * @param link         连接
	 * @return Hyperlink
	 */
	public static Hyperlink getHyperlink(CreationHelper createHelper, String link) {
		Hyperlink hyperlink = null;
		if (link.startsWith(ExcelConstant.HTTPS_PROTOCOL) || link.startsWith(ExcelConstant.HTTP_PROTOCOL)) {
			hyperlink = createHelper.createHyperlink(HyperlinkType.URL);
		} else if (link.startsWith(ExcelConstant.EMAIL_PROTOCOL)) {
			hyperlink = createHelper.createHyperlink(HyperlinkType.EMAIL);
		} else if (link.startsWith(ExcelConstant.FILE_PROTOCOL)) {
			hyperlink = createHelper.createHyperlink(HyperlinkType.FILE);
		}
		return hyperlink;
	}

	/**
	 * 创建链接
	 *
	 * @param cell
	 * @param value
	 * @param linkName
	 * @param createHelper
	 * @return
	 */
	public static Object createHyperlink(Cell cell, Object value, String linkName, CreationHelper createHelper) {
		if (StringUtil.notEmpty(linkName)) {
			String address = value.toString();
			Hyperlink hyperlink = getHyperlink(createHelper, address);
			if (null != hyperlink) {
				hyperlink.setAddress(address);
				cell.setHyperlink(hyperlink);
			}
			if (!linkName.equals(ExcelConstant.MINUS_ONE_STR)) {
				value = linkName;
			}
		}
		return value;
	}

	/**
	 * 设置单元格值
	 *
	 * @param cell
	 * @param value
	 */
	public static void setCellValue(Cell cell, Object value) {
		// 设置内容
		if (value instanceof String) {
			cell.setCellValue((String) value);
		} else if (value instanceof Number) {
			Number number = (Number) value;
			cell.setCellValue(number.doubleValue());
		} else if (value instanceof Boolean) {
			cell.setCellValue((boolean) value);
		} else if (value instanceof Date) {
			cell.setCellValue((Date) value);
		} else if (value instanceof Calendar) {
			cell.setCellValue((Calendar) value);
		} else if (value instanceof RichTextString) {
			cell.setCellValue((RichTextString) value);
		} else {
			cell.setCellValue(value.toString());
		}
	}

	/**
	 * 根据名称获取sheet，不存在则创建
	 *
	 * @param workbook workbook
	 * @param name     名称
	 * @return Sheet
	 */
	public static Sheet getSheetOrCreate(Workbook workbook, String name) {
		Sheet sheet = workbook.getSheet(name);
		if (Objects.isNull(sheet)) {
			sheet = workbook.createSheet(name);
		}
		return sheet;
	}

	/**
	 * 获取行，不存在则创建
	 *
	 * @param sheet    sheet对象
	 * @param rowIndex 行坐标
	 * @return Row
	 */
	public static Row getRowOrCreate(Sheet sheet, int rowIndex) {
		synchronized (sheet) {
			Row row = sheet.getRow(rowIndex);
			if (Objects.isNull(row)) {
				row = sheet.createRow(rowIndex);
			}
			return row;
		}
	}

	/**
	 * 获取单元格，不存在则创建
	 *
	 * @param row       行
	 * @param cellIndex 单元格坐标
	 * @return Cell
	 */
	public static Cell getCellOrCreate(Row row, int cellIndex) {
		Cell cell = row.getCell(cellIndex);
		if (Objects.isNull(cell)) {
			cell = row.createCell(cellIndex);
		}
		return cell;
	}

	/**
	 * 根据rowIndex、cellIndex获取cell
	 * @param sheet - sheet
	 * @param rowIndex - rowIndex
	 * @param cellIndex - cellIndex
	 * @return
	 */
	public static Cell getCellOrCreate(Sheet sheet, int rowIndex,int cellIndex) {
		Row row = getRowOrCreate(sheet, rowIndex);
		return getCellOrCreate(row, cellIndex);
	}

	/**
	 * 设置行高
	 *
	 * @param row       行
	 * @param rowHeight 行高
	 */
	public static void setRowHeight(Row row, short rowHeight) {
		if (ExcelConstant.MINUS_TWO_SHORT != rowHeight) {
			row.setHeightInPoints(rowHeight);
		}
	}

	/**
	 * 设置列宽
	 *
	 * @param sheet     sheet
	 * @param cellIndex 列坐标
	 * @param width     宽度
	 */
	public static void setColWidth(Sheet sheet, int cellIndex, short width) {
		if (ExcelConstant.MINUS_TWO_SHORT != width) {
			int widthPixel = (int) ExcelConstant.PIXEL_RATE * width;
			if (ExcelConstant.MINUS_ONE_SHORT == width) {
				sheet.autoSizeColumn(cellIndex);
			} else {
				sheet.setColumnWidth(cellIndex, widthPixel);
			}
		}
	}

	/**
	 * 移动row
	 *
	 * @param sheet
	 * @param rowIndex
	 * @return
	 */
	public static void shiftRows(Sheet sheet, Integer rowIndex, int shiftRowSize) {
		if (sheet.getRow(rowIndex) != null) {
			int lastRowNo = sheet.getLastRowNum();
			sheet.shiftRows(rowIndex, lastRowNo, shiftRowSize);
		}
	}

	public static void addDropDownValidation(Sheet sheet, CellRangeAddressList regions, String[] options) {
		addDropDownValidation(sheet, regions, options == null ? Collections.emptyList() : Arrays.asList(options), null, null);
	}

	public static void addRangeValidation(Sheet sheet, CellRangeAddressList regions, String start, String end) {
		addRangeValidation(sheet, regions, start, end, null, null);
	}

	/**
	 * 获取总行数
	 *
	 * @param workbook
	 * @param sheetName
	 * @return
	 */
	public static int getRowNum(Workbook workbook, String sheetName) {
		Sheet sheet = workbook.getSheet(sheetName);
		if (Objects.isNull(sheet)) {
			return ExcelConstant.ZERO_SHORT;
		}
		return sheet.getLastRowNum() + ExcelConstant.ONE_INT;
	}

	/**
	 * 获取总列数
	 *
	 * @param workbook
	 * @param sheetName
	 * @return
	 */
	public static int getCellNum(Workbook workbook, String sheetName, int rowIndex) {
		Sheet sheet = workbook.getSheet(sheetName);
		if (Objects.isNull(sheet)) {
			return ExcelConstant.ZERO_SHORT;
		}
		Row row = sheet.getRow(rowIndex);
		if (Objects.isNull(row)) {
			return ExcelConstant.ZERO_SHORT;
		}
		return row.getPhysicalNumberOfCells();
	}

	/**
	 * 添加下拉框数据验证
	 *
	 * @param sheet
	 * @param addressList
	 * @param dropDownList
	 */
	public static void addDropDownValidation(Sheet sheet, CellRangeAddressList addressList, List<String> dropDownList,String title,String message) {
		DataValidationHelper dvHelper = sheet.getDataValidationHelper();
		DataValidationConstraint dvConstraint = dvHelper.createExplicitListConstraint(dropDownList.toArray(new String[0]));
		DataValidation validation = dvHelper.createValidation(dvConstraint, addressList);
		if (validation instanceof XSSFDataValidation) {
			validation.setSuppressDropDownArrow(true);
			validation.setShowErrorBox(true);
		} else {
			validation.setSuppressDropDownArrow(false);
		}
		validation.setErrorStyle(DataValidation.ErrorStyle.STOP);
		if (StringUtil.isEmpty(title)) {
			title = ExcelErrorMsgConstant.ERROR_DROP_DOWN_TITLE;
		}
		if (StringUtil.isEmpty(message)) {
			message = ExcelErrorMsgConstant.ERROR_DROP_DOWN_MSG;
		}
		validation.createErrorBox(title, message);
		sheet.addValidationData(validation);
	}

	/**
	 * 添加范围验证
	 *
	 * @param sheet
	 * @param addressList
	 * @param startNumStr
	 * @param endNumStr
	 */
	public static void addRangeValidation(Sheet sheet, CellRangeAddressList addressList, String startNumStr, String endNumStr,String title,String message) {
		DataValidationHelper dvHelper = sheet.getDataValidationHelper();
		DataValidationConstraint dvConstraint = dvHelper.createNumericConstraint(DVConstraint.ValidationType.INTEGER, DVConstraint.OperatorType.BETWEEN, startNumStr, endNumStr);
		DataValidation validation = dvHelper.createValidation(dvConstraint, addressList);
		if (validation instanceof XSSFDataValidation) {
			validation.setSuppressDropDownArrow(true);
			validation.setShowErrorBox(true);
		} else {
			validation.setSuppressDropDownArrow(false);
		}
		if (StringUtil.isEmpty(title)) {
			title = ExcelErrorMsgConstant.ERROR_DROP_DOWN_TITLE;
		}
		if (StringUtil.isEmpty(message)) {
			message = String.format(ExcelErrorMsgConstant.ERROR_RANGE_MSG, startNumStr, endNumStr);
		}
		validation.setErrorStyle(DataValidation.ErrorStyle.STOP);
		validation.createErrorBox(title, message);
		sheet.addValidationData(validation);
	}

	/**
	 * 创建命名
	 *
	 * @param wb
	 * @param namename
	 * @param formula
	 */
	public static void addRefersToFormula(Workbook wb, String nameName, String formula) {
		Name name = wb.createName();
		name.setNameName(nameName);
		name.setRefersToFormula(formula);
	}

	/**
	 * 创建批注
	 *
	 * @param wb
	 * @param sheet
	 * @param rowIndex
	 * @param colIndex
	 */
	public static void createComment(Workbook wb, Sheet sheet, int rowIndex, int colIndex, String author, RichTextString richTextString, CreationHelper creationHelper) {
		if (Objects.isNull(wb)) {
			throw new ExcelWriterException("Workbook can't be null");
		}
		if (Objects.isNull(sheet)) {
			throw new ExcelWriterException("sheet can't be null");
		}
		Drawing drawing = sheet.createDrawingPatriarch();

		Row row = getRowOrCreate(sheet, rowIndex);
		Cell cell = getCellOrCreate(row, colIndex);

		// When the comment box is visible, have it show in a 1x3 space
		ClientAnchor anchor = creationHelper.createClientAnchor();
		anchor.setCol1(cell.getColumnIndex());
		anchor.setCol2(cell.getColumnIndex() + ExcelConstant.THREE_INT);
		anchor.setRow1(row.getRowNum());
		anchor.setRow2(row.getRowNum() + ExcelConstant.THREE_INT);

		// Create the comment and set the text+author
		Comment comment = drawing.createCellComment(anchor);
		comment.setString(richTextString);
		comment.setAuthor(author);
		// Assign the comment to the cell
		cell.setCellComment(comment);
	}

	public static void createComment(Workbook wb, Sheet sheet, int rowIndex, int colIndex, String author, String commentText, Font font) {
		if (Objects.isNull(commentText)) {
			return;
		}
		CreationHelper creationHelper = wb.getCreationHelper();
		RichTextString richTextString = creationHelper.createRichTextString(commentText);
		if (font != null) {
			richTextString.applyFont(font);
		}
		createComment(wb, sheet, rowIndex, colIndex, author, richTextString, creationHelper);
	}

	/**
	 * 设置sheet 缩放率
	 * @param wb workbook
	 * @param sheetName sheetName
	 * @param scale scale
	 */
	public static void setSheetZoom(Workbook wb, String sheetName,int scale) {
		Sheet sheet = wb.getSheet(sheetName);
		if (Objects.nonNull(sheet)) {
			sheet.setZoom(scale);
		}
	}

	/**
	 * 设置打印区域
	 * @param wb
	 * @param sheetIndex
	 * @param startColIndex
	 * @param endColIndex
	 * @param startRowIndex
	 * @param endRowIndex
	 */
	public static void setPrintArea(Workbook wb, int sheetIndex,int startColIndex,int endColIndex,int startRowIndex,int endRowIndex) {
		wb.setPrintArea(sheetIndex,startColIndex,endColIndex,startRowIndex,endRowIndex);
	}

	/**
	 * 设置脚部页码
	 * @param wb
	 * @param sheetName
	 */
	public static void setFooterNumberByDefault(Workbook wb, String sheetName) {
		Sheet sheet = wb.getSheet(sheetName);
		if (Objects.nonNull(sheet)) {
			Footer footer = sheet.getFooter();
			footer.setCenter(String.format(ExcelConstant.STRING_DEFAULT_FOOTER_TEXT, HeaderFooter.page(),HeaderFooter.numPages()));
		}
	}

	/**
	 * 设置脚部页码
	 * @param wb
	 * @param sheetName
	 */
	public static void setFooterNumber(Workbook wb, String sheetName,String formatPattern) {
		Sheet sheet = wb.getSheet(sheetName);
		if (Objects.nonNull(sheet)) {
			Footer footer = sheet.getFooter();
			footer.setCenter(String.format(formatPattern, HeaderFooter.page(),HeaderFooter.numPages()));
		}
	}

	/**
	 * 创建拆分窗格
	 * @param wb workbook
	 * @param sheetName sheetName
	 * @param xSplitPos x 轴坐标，等同于像素
	 * @param ySplitPos y 坐标，等同于像素
	 * @param leftmostColumn 左边单元格数量
	 * @param topRow 上面单元格数量
	 */
	public static void createSplitPane(Workbook wb, String sheetName,int xSplitPos, int ySplitPos, int leftmostColumn, int topRow) {
		Sheet sheet = wb.getSheet(sheetName);
		if (Objects.nonNull(sheet)) {
			sheet.createSplitPane(xSplitPos,ySplitPos,leftmostColumn,topRow,Sheet.PANE_LOWER_LEFT);
		}
	}

	/**
	 * 自定义颜色
	 * @param workbook
	 * @param r
	 * @param g
	 * @param b
	 * @return
	 */
	public static HSSFColor setCustomColor(HSSFWorkbook workbook, byte r, byte g, byte b,short colorIndex){
		HSSFPalette palette = workbook.getCustomPalette();
		HSSFColor hssfColor;
		try {
			hssfColor= palette.findColor(r, g, b);

			if (hssfColor == null ){
				palette.setColorAtIndex(colorIndex, r, g,b);
				hssfColor = palette.getColor(colorIndex);
			}
		} catch (Exception e) {
			throw new ExcelWriterException(e.getMessage());
		}
		return hssfColor;
	}
	/**
	 * @Description: 获取单元格值
	 * @Param:
	 * @return:
	 * @Author: Vachel Wang
	 * @Date: 2026/4/24
	 */
	public static Object getCellValue(Cell cell , Class cla, String colAddress,FormulaEvaluator evaluator) {
		Object valueObj = null;
		if (Objects.isNull(cell)) return valueObj;
		switch (cell.getCellType()) {
			case STRING:
				valueObj = getCellNumberValue(cell, cla, colAddress);
				break;
			case NUMERIC:
				if (DateUtil.isCellDateFormatted(cell)) {
					valueObj = cell.getDateCellValue();
				} else {
					if(cla == String.class){
						valueObj = getCellNumberValue(cell, cla, colAddress);
					}else{
						valueObj = cell.getNumericCellValue();
					}
				}
				break;
			case BOOLEAN:
				valueObj = String.valueOf(cell.getBooleanCellValue());
				break;
			case FORMULA:
				evaluator.evaluateInCell(cell);
				CellValue cellValue = evaluator.evaluate(cell);
				if (StringUtil.notEmpty(cellValue.getStringValue())) {
					valueObj = cellValue.getStringValue();
				} else {
					valueObj = cellValue.getNumberValue();
				}
				break;
			case ERROR:
				valueObj = null ;
			default:
				valueObj = null;
		}
		return valueObj;
	}

	private static Object getCellNumberValue(Cell cell, Class cla, String colAddress) {
		Object valueObj ;
		if(cell.getCellType() == CellType.STRING){
			valueObj = cell.getRichStringCellValue().getString();
		} else if (cell.getCellType() == CellType.NUMERIC) {
			cell.getCellStyle().setDataFormat(cell.getRow().getSheet().getWorkbook().getCreationHelper().createDataFormat().getFormat(ExcelConstant.NUMBER_PATTERN_FORMAT));
			valueObj = DATA_FORMATTER.formatCellValue(cell);
		}else{
			valueObj = DATA_FORMATTER.formatCellValue(cell);
		}
		if(Number.class.isAssignableFrom(cla)){
			// 正则校验数字是否符合规范
			if(!NUMBER_PATTERN.matcher(valueObj.toString()).matches()) {
				throw new ExcelReaderException("单元格：[" + colAddress + "]需要填写数字，当前值：" + valueObj);
			}
		}
		if (StringUtil.isEmpty(valueObj)) {
			valueObj = null;
		}
		return valueObj;
	}

	/**
	 * @Description: 获取图片内容
	 * @Param:
	 * @return:
	 * @Author: Vachel Wang
	 * @Date: 2026/4/24
	 */
	public static Object getPictureValue(Map<String, List<ExcelReaderPictureModel>> sheetPictureMap, int rowIndex, int colIndex, Object setParams, Class<?> parameterType) {
		if(Objects.isNull(sheetPictureMap) || sheetPictureMap.isEmpty()) return setParams;
		if (ExcelReaderPictureModel.class.isAssignableFrom(parameterType)) {
			List<ExcelReaderPictureModel> pictureModelList = sheetPictureMap.get(String.valueOf(rowIndex) + ExcelConstant.COMMA_CHAR + colIndex);
			if (!pictureModelList.isEmpty()) {
				setParams = pictureModelList.get(ExcelConstant.ZERO_SHORT);
			}
		} else if (List.class.isAssignableFrom(parameterType)) {
			setParams = sheetPictureMap.get(String.valueOf(rowIndex) + ExcelConstant.COMMA_CHAR + colIndex);
		}else if(parameterType.getTypeName().equals(ExcelConstant.BYTE_ARRAY_STR)){
			List<ExcelReaderPictureModel> pictureModelList = sheetPictureMap.get(String.valueOf(rowIndex) + ExcelConstant.COMMA_CHAR + colIndex);
			if (Objects.nonNull(pictureModelList) && !pictureModelList.isEmpty()) {
				setParams = pictureModelList.get(ExcelConstant.ZERO_SHORT).getBytes();
			}
		}
		return setParams;
	}
}
