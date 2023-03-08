package com.excel.workbook.common.model;

import com.excel.annotation.ExcelHeader;
import lombok.Data;

@Data
public class CommonModel {

    @ExcelHeader(headerName = "no", priority = 100)
    private String id;
}
