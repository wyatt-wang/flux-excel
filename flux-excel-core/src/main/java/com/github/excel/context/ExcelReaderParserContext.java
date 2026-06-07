package com.github.excel.context;

import com.github.excel.model.ExcelBaseModel;
import com.github.excel.model.ExcelReaderPictureModel;
import com.github.excel.read.executor.ExcelReaderPictureExecutor;
import com.github.excel.read.executor.ExcelReaderTemplateValidator;
import com.github.excel.read.config.ReadModelTitleConfig;
import com.github.excel.read.format.ExcelReaderFormatManager;
import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;
import lombok.experimental.Accessors;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

import java.util.List;
import java.util.Map;

/**
 * excel解析上下文
 * @author Vico
 * @create 2023-08-18 18:24
 */
@Data
@Accessors(chain = true)
@Builder
@AllArgsConstructor
@NoArgsConstructor
public class ExcelReaderParserContext<T extends ExcelBaseModel> {
    /**
     * 图片执行
     */
    private ExcelReaderPictureExecutor pictureExecutor;
    /**
     * 模版校验
     */
    private ExcelReaderTemplateValidator templateValidator;
    /**
     * 格式化器
     */
    private ExcelReaderFormatManager readerFormatManager;
    /**
     * sheet 解析上下文
     */
    private List<ExcelReaderModelContext<T>> readModelList;
    /**
     * 图片map
     */
    private Map<String, List<ExcelReaderPictureModel>> sheetPictureMap;
    /**
     * 行
     */
    private Row row;
    /**
     * 列
     */
    private Cell cell;
    /**
     * 值
     */
    private Object cellValue;
    /**
     * 填充model列表
     */
    private List<ExcelReaderModelContext<T>> fillModels ;
    /**
     * 填充list列表
     */
    private List<ExcelReaderModelContext<T>> fillModelLists ;
    /**
     * 标题配置
     */
    private ReadModelTitleConfig<T> listTitleConfig;
    /**
     * 当前modelContext
     */
    private ExcelReaderModelContext<T> currentModelContext;
}
