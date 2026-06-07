package com.github.excel.helper;

import com.github.excel.constant.ExcelConstant;
import com.github.excel.model.ExcelBaseModel;
import org.apache.commons.collections4.CollectionUtils;

import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Objects;

/**
 * 导出帮助类
 * @author Vico
 * @create 2023-09-08 13:55
 */
public class ExcelWriterHelper {

    /**
     * 排除
     *
     * @param excludeFields
     */
    public static void excludes(Class<? extends ExcelBaseModel> modelCla, List<String> excludeFields , Map<Class<? extends ExcelBaseModel>, Map<String, String>> excludeFieldMap) {
        if (CollectionUtils.isNotEmpty(excludeFields)) {
            Map<String, String> fieldMap = excludeFieldMap.get(modelCla);
            boolean exists = true;
            if (Objects.isNull(fieldMap)) {
                fieldMap = new HashMap<>();
                exists = false;
            }
            for (String field : excludeFields) {
                fieldMap.put(field, ExcelConstant.NULL_STR);
            }
            if (!exists) {
                excludeFieldMap.put(modelCla, fieldMap);
            }
        }
    }

    /**
     * 添加验证或批注
     *
     * @param field
     * @param excelWriterCommentParam
     * @return
     */
    /*public ExcelWriterExcludeHandler addValidationOrComment(String field, ExcelWriterCommentParam excelWriterCommentParam) {
        if (Objects.isNull(excelWriterCommentParam) || StringUtil.isEmpty(field)) {
            return this;
        }
        Map<String, ExcelWriterCommentParam> fieldMap = commentMap.get(modelCla);
        boolean exists = true;
        if (Objects.isNull(fieldMap)) {
            fieldMap = new HashMap<>();
            exists = false;
        }
        fieldMap.put(field, excelWriterCommentParam);
        if (!exists) {
            commentMap.put(modelCla, fieldMap);
        }
        return this;
    }*/
}
