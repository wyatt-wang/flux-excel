package com.github.excel.config;

import lombok.Data;
import org.springframework.boot.context.properties.ConfigurationProperties;
import org.springframework.stereotype.Component;

import java.util.List;
import java.util.Map;
import java.util.concurrent.ConcurrentHashMap;

/**
 * @Author: Vachel Wang
 * @Date: 2026/4/24
 * @Description: flux-excel配置
 */
@Data
@ConfigurationProperties(prefix = "excel")
@Component
public class ExcelProperties {

    public static final Map<String, String> configMap = new ConcurrentHashMap<>();
    /**
     * 包路径，已废弃。模型元数据默认按使用时传入的 Class 懒加载。
     */
    @Deprecated
    private List<String> dtoPackage;
    /**
     * 导出模板文件夹
     */
    private String exportTemplateDir;
    /**
     * 导入模板文件夹
     */
    private String importTemplateDir;
}
