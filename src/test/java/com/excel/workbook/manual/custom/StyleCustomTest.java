package com.excel.workbook.manual.custom;

import com.excel.core.CstWorkBook;
import com.excel.core.CstWorkSheet;
import com.excel.core.type.CstWorkBookType;
import com.excel.core.type.SheetDirection;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.junit.jupiter.api.Assertions;
import org.junit.jupiter.api.Test;

import java.io.IOException;
import java.util.stream.IntStream;

public class StyleCustomTest {

    @Test
    public void test_01() throws IOException {
        CstWorkBook workBook = new CstWorkBook(CstWorkBookType.XSSF);
        CstWorkSheet sheet = workBook.createSheet().updateDirection(SheetDirection.VERTICAL);

        sheet.createTitleCell(2, "이름", "나이");
        sheet.newLine();


        IntStream.range(0, 2).forEach(i -> {
            sheet.createCell( "희섭" + i, 1 + i);
            sheet.newLine();
        });

        Cell cell = sheet.getCell(0, 1);

        Assertions.assertTrue(cell.getStringCellValue().equals("희섭0"));
        Assertions.assertTrue(cell.getCellStyle().getFillForegroundColor() == IndexedColors.BLUE.getIndex());

        workBook.write("target/excel/manual/custom/custom_style_01.xls");


    }
}
