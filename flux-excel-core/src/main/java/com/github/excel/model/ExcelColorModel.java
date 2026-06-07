package com.github.excel.model;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.awt.*;

/**
 * 颜色model
 * @author Vico
 * @create 2024-08-07 下午3:19
 */
@Data
@AllArgsConstructor
@NoArgsConstructor
public class ExcelColorModel {
    private Color color;
    private int index ;
}
