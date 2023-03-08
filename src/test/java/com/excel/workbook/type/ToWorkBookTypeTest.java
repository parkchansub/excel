package com.excel.workbook.type;

import com.excel.core.type.CstWorkBookType;
import org.junit.jupiter.api.Assertions;
import org.junit.jupiter.api.Test;

public class ToWorkBookTypeTest {

    @Test
    public void translate_hssf_ext_test(){
        String ext = CstWorkBookType.HSSF.translateFileName("user.xl");
        Assertions.assertTrue(ext.equals("user.xls"));
    }

    @Test
    public void translate_xssf_ext_test(){
        String ext = CstWorkBookType.XSSF.translateFileName("user.xl");
        Assertions.assertTrue(ext.equals("user.xlsx"));
    }
}
