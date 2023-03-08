package com.excel.core;

import com.excel.core.type.CstWorkBookType;
import com.excel.exception.SheetNotFoundException;
import lombok.NonNull;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.*;
import java.nio.file.Files;
import java.util.ArrayList;
import java.util.List;
import java.util.stream.Collectors;
import java.util.stream.IntStream;

public class CstWorkBook {

    private final Workbook workbook;

    private final CstWorkBookType type;


    private List<CstWorkSheet> sheets = new ArrayList<>();

    public CstWorkBook(CstWorkBookType type) {
        this.type = type;
        this.workbook = type.createWorkBookInstance();
    }


    /**
     * Instantiates a new To work book.
     *
     * @param file the file
     * @throws IOException the io exception
     */
    public CstWorkBook(@NonNull File file) throws IOException {
        this(CstWorkBookType.findWorkBookType(file.getName()), file);
    }

    /**
     * Instantiates a new To work book.
     *
     * @param type the type
     * @param file the file
     * @throws IOException the io exception
     */
    public CstWorkBook(@NonNull CstWorkBookType type, @NonNull File file) throws IOException {
        if (!file.exists() || file.isDirectory()) {
            throw new FileNotFoundException();
        }

        this.type = type;
        this.workbook = type.createWorkBookInstance(file);
        this.__initSheet();
    }

    /**
     * Instantiates a new To work book.
     *
     * @param type the type
     * @param fis  the fis
     * @throws IOException the io exception
     */
    public CstWorkBook(@NonNull CstWorkBookType type, @NonNull FileInputStream fis) throws IOException {
        this.type = type;
        this.workbook = type.createWorkBookInstance(fis);
        this.__initSheet();
    }

    /**
     * Create sheet to work sheet.
     *
     * @return the to work sheet
     */
    public CstWorkSheet createSheet() {
        CstWorkSheet sheet = new CstWorkSheet(this.workbook);
        this.sheets.add(sheet);
        return sheet;
    }

    /**
     * Create sheet to work sheet.
     *
     * @param name the name
     * @return the to work sheet
     */
    public CstWorkSheet createSheet(String name) {
        CstWorkSheet sheet = new CstWorkSheet(this.workbook, name);
        this.sheets.add(sheet);
        return sheet;
    }


    /**
     * Write.
     *
     * @param filePath the file path
     * @throws IOException the io exception
     */
    public void write(@NonNull String filePath) throws IOException {
        String fp = this.type.translateFileName(filePath);
        File file = new File(fp);
        Files.createDirectories(file.getParentFile().toPath());

        FileOutputStream fileOutputStream = new FileOutputStream(fp);
        this.workbook.write(fileOutputStream);
    }

    public String getFileName(String name){
        if( this.type == null) { return name; }
        return type.translateFileName(name);
    }

    /**
     * Write.
     *
     * @param outputStream the output stream
     * @throws IOException the io exception
     */
    public void write(@NonNull OutputStream outputStream) throws IOException {
        this.workbook.write(outputStream);
    }

    /**
     * Sheet size int.
     *
     * @return the int
     */
    public int sheetSize() {
        return this.sheets.size();
    }

    /**
     * Gets sheet at.
     *
     * @param idx the idx
     * @return the sheet at
     */
    public CstWorkSheet getSheetAt(int idx) {
        return this.sheets.get(idx);
    }

    /**
     * Gets sheet.
     *
     * @param name the name
     * @return the sheet
     */
    public CstWorkSheet getSheet(@NonNull String name) {
        return this.sheets.stream()
                .filter(st -> name.equalsIgnoreCase(st.getName()))
                .findFirst().orElseThrow(() -> new SheetNotFoundException("Not found sheet " + name));
    }



    private void __initSheet() {
        this.sheets = IntStream.range(0, this.workbook.getNumberOfSheets()).mapToObj(this.workbook::getSheetAt)
                .map(sheet -> new CstWorkSheet( sheet)).collect(Collectors.toList());
    }


}
