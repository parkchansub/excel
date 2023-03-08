package com.excel.workbook.mapping;

import com.excel.core.CstWorkBook;
import com.excel.core.CstWorkSheet;
import com.excel.core.type.CstWorkBookType;
import com.excel.workbook.common.model.UserModel;
import org.junit.jupiter.api.Assertions;
import org.junit.jupiter.api.Test;

import java.io.IOException;
import java.util.List;
import java.util.stream.Collectors;
import java.util.stream.IntStream;

public class AutoMapptingFromModelTest {

    /*@Test
    public void userModelsToExcelWriteTest_01() throws IOException {
        List<UserModel> userModelList =
                IntStream.range(0, 100).mapToObj(i ->
                        UserModel.builder().name("tester" + i).age(i).gender("man").build()).collect(Collectors.toList());

        CstWorkBook workBook = new CstWorkBook(CstWorkBookType.XSSF);
        CstWorkSheet sheet = workBook.createSheet();
        sheet.from(userModelList);

        Assertions.assertTrue(sheet.getCell(0, 0).getStringCellValue().equalsIgnoreCase("name"));
        Assertions.assertTrue(sheet.getCell(0, 1).getStringCellValue().equalsIgnoreCase("age"));
        Assertions.assertTrue(sheet.getCell(0, 2).getStringCellValue().equalsIgnoreCase("gender"));

        Assertions.assertTrue(sheet.getCell(1, 0).getStringCellValue().equalsIgnoreCase("tester0"));
        Assertions.assertTrue(sheet.getCell(1, 1).getNumericCellValue() == 0 );
        Assertions.assertTrue(sheet.getCell(1, 2).getStringCellValue().equalsIgnoreCase("man"));
        Assertions.assertTrue(sheet.getRowCount() == 101);


        workBook.write("target/excel/map/write_test_1");

    }*/
}
