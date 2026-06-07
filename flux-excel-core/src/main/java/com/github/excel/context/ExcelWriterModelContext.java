package com.github.excel.context;

import com.github.excel.model.ExcelCacheModel;
import com.github.excel.param.ExcelWriterDataParam;
import com.github.excel.param.ExcelWriterParam;
import com.github.excel.write.ExcelCustomWriter;
import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;
import lombok.experimental.Accessors;

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
public class ExcelWriterModelContext {
    /**
     * 写入bean或list参数
     */
    private ExcelWriterDataParam writerDataParam;
    /**
     * 缓存模型
     */
    private ExcelCacheModel cacheModel ;
}
