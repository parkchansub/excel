package com.excel.core.type;

import com.excel.exception.NotFoundExtException;
import lombok.NonNull;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.Arrays;

/**
 * The enum Work book type.
 */
public enum CstWorkBookType {

    /**
     * Hssf work book type.
     */
    HSSF("xls"),
    /**
     * Xssf work book type.
     */
    XSSF("xlsx");

    private final String ext;

    CstWorkBookType(String ext) {
        this.ext = ext;
    }

    /**
     * Create work book instance workbook.
     *
     * @return the workbook
     */
    public Workbook createWorkBookInstance() {
        if (this.equals(HSSF)) {
            /**
             *  excel 97 ~ 2003
             */
            return new HSSFWorkbook();
        }
        /**
         *  excel 2007 이상
         */
        return new XSSFWorkbook();

    }

    /**
     * Find work book type work book type.
     *
     * @param ext the ext
     * @return the work book type
     */
    public static CstWorkBookType findWorkBookType(@NonNull String ext) {
        return Arrays.stream(CstWorkBookType.values()).filter(type -> ext.endsWith(type.ext)).findFirst()
                .orElseThrow(() -> new NotFoundExtException(ext + " is not supported"));
    }

    /**
     * Create work book instance workbook.
     *
     * @param file the file
     * @return the workbook
     * @throws IOException the io exception
     */
    public Workbook createWorkBookInstance(@NonNull File file) throws IOException {
        FileInputStream fis = new FileInputStream(file);
        return createWorkBookInstance(fis);
    }

    /**
     * Create work book instance workbook.
     *
     * @param fis the fis
     * @return the workbook
     * @throws IOException the io exception
     */
    public Workbook createWorkBookInstance(@NonNull FileInputStream fis) throws IOException {
        if (this.equals(HSSF)) {
            return new HSSFWorkbook(fis);
        }
        return new XSSFWorkbook(fis);
    }

    /**
     * Translate file name string.
     *
     * @param filename the filename
     * @return the string
     */
    public String translateFileName(@NonNull String filename) {
        String _fp = filename;
        if (_fp.lastIndexOf(".") > 0) {
            _fp = _fp.substring(0, _fp.indexOf("."));
        }
        return _fp + "." + this.ext;
    }


}
