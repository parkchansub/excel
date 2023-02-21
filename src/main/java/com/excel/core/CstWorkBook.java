package com.excel.core;

import com.excel.core.type.ToWorkBookType;
import org.apache.poi.ss.usermodel.Workbook;

import java.util.ArrayList;
import java.util.List;

public class CstWorkBook {

    private final Workbook workbook;

    private final ToWorkBookType type;


    private List<CstWorkSheet> sheets = new ArrayList<>();

    public CstWorkBook(ToWorkBookType type) {
        this.type = type;
        this.workbook = type.createWorkBookInstance();
    }


}
