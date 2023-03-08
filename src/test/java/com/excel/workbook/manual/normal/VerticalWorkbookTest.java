package com.excel.workbook.manual.normal;

import com.excel.core.CstWorkBook;
import com.excel.core.CstWorkSheet;
import com.excel.core.type.CstWorkBookType;
import com.excel.core.type.SheetDirection;
import org.junit.jupiter.api.Assertions;
import org.junit.jupiter.api.Test;

import java.io.IOException;
import java.util.stream.IntStream;

public class VerticalWorkbookTest {

    @Test
    public void vertical_test_01() throws IOException {
        CstWorkBook workBook = new CstWorkBook(CstWorkBookType.XSSF);
        CstWorkSheet sheet = workBook.createSheet().updateDirection(SheetDirection.VERTICAL);

        sheet.createTitleCell(2, "name", "age");
        IntStream.range(0, 2).forEach(i ->{
            sheet.createCellToNewline("hsim" + i , 1 + i);
        });

        Assertions.assertTrue(sheet.getCell(0, 0).getStringCellValue().equals("name"));
        Assertions.assertTrue(sheet.getCell(1, 0).getStringCellValue().equals("age"));

        Assertions.assertTrue(sheet.getCell(0, 1).getStringCellValue().equals("hsim0"));
        Assertions.assertTrue(sheet.getCell(1, 1).getNumericCellValue() == 1 ) ;

        workBook.write("target/excel/manual/normal/vertical_test_01");


    }
}
