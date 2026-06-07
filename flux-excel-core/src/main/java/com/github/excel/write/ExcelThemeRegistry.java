package com.github.excel.write;

import com.github.excel.enums.ExcelThemeEnum;
import com.google.common.collect.Maps;

import java.util.Map;
import java.util.Objects;

/**
 * Runtime registry for named Excel themes.
 */
public final class ExcelThemeRegistry {
    private static final Map<String, ExcelTheme> THEME_MAP = Maps.newConcurrentMap();

    static {
        for (ExcelThemeEnum value : ExcelThemeEnum.values()) {
            register(value.name(), value);
        }
    }

    private ExcelThemeRegistry() {
    }

    public static void register(String name, ExcelTheme theme) {
        if (name == null || theme == null) {
            return;
        }
        THEME_MAP.put(name, theme);
    }

    public static ExcelTheme getTheme(String name) {
        if (name == null || name.trim().isEmpty()) {
            return ExcelThemeEnum.NONE;
        }
        ExcelTheme theme = THEME_MAP.get(name);
        return Objects.nonNull(theme) ? theme : ExcelThemeEnum.NONE;
    }
}
