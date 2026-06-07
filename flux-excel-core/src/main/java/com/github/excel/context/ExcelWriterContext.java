package com.github.excel.context;

import com.github.excel.model.*;
import com.github.excel.param.ExcelWriterParam;
import com.github.excel.write.*;
import com.github.excel.write.style.AbstractExcelStyle;
import com.github.excel.write.style.ExcelBasicStyle;
import com.google.common.collect.Maps;
import com.google.common.collect.Sets;
import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;
import lombok.experimental.Accessors;
import org.apache.poi.ss.usermodel.*;

import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.concurrent.ConcurrentHashMap;

/**
 * Excel写，上下文
 * @author Vico
 * @create 2023-08-17 14:54
 */
@Data
@Accessors(chain = true)
@Builder
@NoArgsConstructor
@AllArgsConstructor
public class ExcelWriterContext {
    /**
     * 自定义写
     */
    private ExcelCustomWriter customWrite;
    /**
     * Excel 模型写入参数
     */
    private ExcelWriterModelContext modelContext = new ExcelWriterModelContext();
    /**
     * Excel 创建参数
     */
    private ExcelWriterParam writerParam;
    /**
     * 工作簿
     */
    private Workbook workbook;
    /**
     * 创建帮助类
     */
    private CreationHelper creationHelper;
    /**
     * 自定义style
     */
    private Set<Class<? extends AbstractExcelStyle>> styles = Sets.newHashSet(ExcelBasicStyle.class);
    /**
     * 样式缓存
     */
    private ThreadLocal<Map<String, CellStyle>> styleLocal = ThreadLocal.withInitial(ConcurrentHashMap::new);
    /**
     * 字体缓存
     */
    private ThreadLocal<Map<String, Font>> fontLocal = ThreadLocal.withInitial(ConcurrentHashMap::new);
    /**
     * 颜色缓存
     */
    private ThreadLocal<Map<String, Color>> colorLocal = ThreadLocal.withInitial(ConcurrentHashMap::new);
    /**
     * 表达式map
     */
    private Map<String, List<ExcelExpressionModel>> templateExpressionMap;
    /**
     * 模版title map
     */
    private Map<Integer, List<ExcelTemplateTitleModel>> templateTitleMap;
    /**
     * 排除map
     */
    private Map<Class<? extends ExcelBaseModel>, Map<String, String>> excludeFieldMap = new HashMap<>();
    /**
     * 下标model
     */
    private ExcelWriterPointModel pointModel;
    /**
     * sheet 会实时变动，线程不安全
     */
    private Sheet sheet ;
    /**
     * row
     */
    private Row row ;
    /**
     * cell
     */
    private Cell cell;
    /**
     * 字段
     */
    private ExcelCacheFieldModel fieldModel;
    /**
     * 值会实时变动，线程不安全
     */
    private Object value;
    /**
     * 样式名称
     */
    private String styleName;
    /**
     * 格式化字符串
     */
    private String formatPattern;
    /**
     * 格式化class
     */
    private Class<? extends ExcelWriterDataFormat> formatClass;
    /**
     * 图片类型
     */
    private Integer pictureType ;
    /**
     * 字段格式化器缓存
     */
    protected Map<Class<? extends ExcelWriterDataFormat>, ExcelWriterDataFormat> dataFormatMap = new HashMap<Class<? extends ExcelWriterDataFormat>, ExcelWriterDataFormat>(){{
        put(ExcelDefaultWriterDataFormat.class, new ExcelDefaultWriterDataFormat());
    }};


}
