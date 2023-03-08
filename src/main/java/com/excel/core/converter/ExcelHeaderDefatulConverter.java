package com.excel.core.converter;

import com.excel.annotation.ExcelHeader;

/**
 * The type Excel header defatul converter.
 */
public class ExcelHeaderDefatulConverter implements ExcelHeaderConverter {

    @Override
    public String headerKeyConverter(ExcelHeader excelHeader) {
        return excelHeader.headerName();
    }
}
