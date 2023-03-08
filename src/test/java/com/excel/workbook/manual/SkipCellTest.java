package com.excel.workbook.manual;

import com.excel.core.CstWorkBook;
import com.excel.core.CstWorkSheet;
import com.excel.core.type.CstWorkBookType;
import com.excel.core.type.SheetDirection;
import org.junit.jupiter.api.Assertions;
import org.junit.jupiter.api.Test;

import java.io.IOException;
import java.util.stream.IntStream;

public class SkipCellTest {

    @Test
    public void skipCellTest_horizon_01() throws IOException {
        CstWorkBook workBook = new CstWorkBook(CstWorkBookType.XSSF);
        CstWorkSheet sheet = workBook.createSheet().updateDirection(SheetDirection.HORIZON);

        sheet.createTitleCell(2, "이름", "나이");
        sheet.newLine();
        IntStream.range(0, 2).forEach(i -> {
            sheet.skip(1);
            sheet.createCell("희섭" + i, 1 + i);
            sheet.newLine();
        });

        Assertions.assertTrue(sheet.getCell(1, 0).getStringCellValue().isEmpty());
        Assertions.assertTrue(sheet.getCell(2, 0).getStringCellValue().isEmpty());
        Assertions.assertTrue(sheet.getCell(1, 1).getStringCellValue().equals("희섭0"));

        Assertions.assertTrue(sheet.getCell(1, 2) != null);
        Assertions.assertTrue(sheet.getCell(2, 2) != null);


        workBook.write("target/excel/manual/skip/skip_horizon_01.xls");
    }

    @Test
    public void skipCellTest_vertical_01() throws IOException {
        CstWorkBook workBook = new CstWorkBook(CstWorkBookType.XSSF);
        CstWorkSheet sheet = workBook.createSheet().updateDirection(SheetDirection.VERTICAL);

        sheet.createTitleCell(2, "이름", "나이");
        sheet.newLine();
        IntStream.range(0, 2).forEach(i -> {
            sheet.skip(1);
            sheet.createCell("희섭" + i, 1 + i);
            sheet.newLine();
        });

        Assertions.assertTrue(sheet.getCell(0, 1).getStringCellValue().isEmpty());
        Assertions.assertTrue(sheet.getCell(0, 2).getStringCellValue().isEmpty());
        Assertions.assertTrue(sheet.getCell(1, 1).getStringCellValue().equals("희섭0"));
        Assertions.assertTrue(sheet.getCell(1, 2).getStringCellValue().equals("희섭1"));

        Assertions.assertTrue(sheet.getCell(0, 1) != null);
        Assertions.assertTrue(sheet.getCell(0, 2) != null);

        workBook.write("target/excel/manual/skip/skip_vertical_01.xls");
    }

}
