package com.excel.util;

import com.excel.annotation.ExcelHeader;
import com.excel.core.converter.ExcelHeaderConverter;
import lombok.NonNull;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Workbook;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.Objects;
import java.util.stream.Collectors;

/**
 * The interface Excel header helper.
 */
public interface ExcelHeaderHelper {



    /**
     * Header list list.
     *
     * @param header               the header
     * @param excelHeaderConverter the excel header converter
     * @return the list
     */
    default List<String> headerList(@NonNull ExcelHeader header, ExcelHeaderConverter excelHeaderConverter) {
        List<String> headers = new ArrayList<>();
        headers.add(header.headerName());
        if (!header.headerName().equals(excelHeaderConverter.headerKeyConverter(header))) {
            headers.add(excelHeaderConverter.headerKeyConverter(header));
        }
        headers.addAll(Arrays.asList(header.headerNames()));

        return headers.stream().filter(Objects::nonNull).map(String::trim).filter(hd -> !hd.isEmpty()).collect(Collectors.toList());
    }

}
