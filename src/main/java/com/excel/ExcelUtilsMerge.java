package com.excel;

import lombok.Data;

@Data
public class ExcelUtilsMerge {

    private int firstRow;

    private int lastRow;

    private int firstColumn;

    private int lastColumn;

    private Object content;

    private boolean isSetHeaderStyle;
}
