package com.excel.workbook.manual;


import com.excel.core.CstWorkBook;
import com.excel.core.CstWorkSheet;
import com.excel.core.type.CstWorkBookType;
import com.excel.core.type.SheetDirection;
import org.junit.jupiter.api.Assertions;
import org.junit.jupiter.api.Test;

import java.io.IOException;

public class MergeCellTest {

    @Test
    public void mergeCellTest_horizon_01() throws IOException {
        CstWorkBook workBook = new CstWorkBook(CstWorkBookType.XSSF);
        CstWorkSheet sheet = workBook.createSheet().updateDirection(SheetDirection.HORIZON);

        sheet.createTitleCell(2, "name", "age", "contactNumber");
        sheet.merge(2, 1); // 2(width) X 1(height) [][]
        sheet.createCellToNewline("sharky", "36", "010-1234-0000", "02-1111-1234");
        sheet.createCellToNewline("melpis", "36", "010-1111-1234", "02-4221-1234");
        sheet.createCellToNewline("heeseob", "32", "010-0000-1234", "-");
        sheet.createCellToNewline("dongjun", "31", "010-4324-1234", "031-4121-1234");

        Assertions.assertTrue(sheet.getMergedRegions().get(0).getFirstRow() == 0 );
        Assertions.assertTrue(sheet.getMergedRegions().get(0).getLastRow() == 0);
        Assertions.assertTrue(sheet.getMergedRegions().get(0).getFirstColumn() == 2);
        Assertions.assertTrue(sheet.getMergedRegions().get(0).getLastColumn() == 3);


        workBook.write("target/excel/manual/merge/merge_horizon_01");
    }

    @Test
    public void mergeCellTest_horizon_02() throws IOException {
        CstWorkBook workBook = new CstWorkBook(CstWorkBookType.XSSF);
        CstWorkSheet sheet = workBook.createSheet().updateDirection(SheetDirection.VERTICAL);

        sheet.createTitleCell(3, "name");
        sheet.merge(1, 3);
        sheet.newLine();

        sheet.createTitleCell(1, "age");
        sheet.merge(1, 3);
        sheet.newLine();

        sheet.createTitleCell(1, "contactNumber");
        sheet.newLine();

        sheet.createTitleCell(2,"phone", "home", "test");
        sheet.createTitleCell(1, "2020");
        sheet.createTitleCell(1, "12");

        sheet = workBook.createSheet().updateDirection(SheetDirection.HORIZON);
        sheet.createCellToNewline("sharky", "36", "010-1234-0000", "02-1111-1234");
        sheet.createCellToNewline("melpis", "36", "010-1111-1234", "02-4221-1234");
        sheet.createCellToNewline("heeseob", "32", "010-0000-1234", "-");
        sheet.createCellToNewline("dongjun", "31", "010-4324-1234", "031-4121-1234");




        /*Assert.assertTrue(sheet.getMergedRegions().size() == 3);

        Assert.assertTrue(sheet.getMergedRegions().get(0).getFirstRow() == 0 && sheet.getMergedRegions().get(0).getLastRow() == 1);
        Assert.assertTrue(sheet.getMergedRegions().get(0).getFirstColumn() == 0 && sheet.getMergedRegions().get(0).getLastColumn() == 0);
        Assert.assertTrue(sheet.getMergedRegions().get(1).getFirstRow() == 0 && sheet.getMergedRegions().get(1).getLastRow() == 1);
        Assert.assertTrue(sheet.getMergedRegions().get(1).getFirstColumn() == 1 && sheet.getMergedRegions().get(1).getLastColumn() == 1);

        Assert.assertTrue(sheet.getMergedRegions().get(2).getFirstRow() == 0 && sheet.getMergedRegions().get(2).getLastRow() == 0);
        Assert.assertTrue(sheet.getMergedRegions().get(2).getFirstColumn() == 2 && sheet.getMergedRegions().get(2).getLastColumn() == 3);*/

        workBook.write("target/excel/manual/merge/merge_horizon_02");
    }

}
