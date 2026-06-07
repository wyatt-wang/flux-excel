package com.github.excel.write;

public interface ExcelTheme {
    String getTitleRowStyleName();

    String getOddRowStyleName();

    String getEvenRowStyleName();

    String getOddRowStyleDateName();

    String getEvenRowStyleDateName();

    String getTitleBackGroundColorName();

    String getBorderColorName();

    String getOddRowBackGroundColorName();

    String getEventRowBackGroundColorName();

    String getTitleFontName();

    String getContentFontName();

    Short getTitleRowHeight();

    Short getContentRowHeight();

    Short getColWidth();
}
