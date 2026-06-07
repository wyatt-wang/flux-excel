package com.github.excel.annotation;

import java.lang.annotation.*;

/**
 * Excel 校验，快速失败
 */
@Documented
@Inherited
@Retention(RetentionPolicy.RUNTIME)
@Target(ElementType.TYPE)
public @interface ExcelValidation {

}
