package com.excel.core;

import com.excel.core.style.CstWorkBookStyle;
import com.excel.core.type.CstWorkCellType;
import lombok.NonNull;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.RichTextString;

import java.util.Calendar;
import java.util.Date;

public class CstWorkCell {

    private final Cell cell;

    private CstWorkCellType cellType;

    /**
     * Instantiates a new To work cell.
     *
     * @param cell  the cell
     * @param value the value
     */
    public CstWorkCell(@NonNull Cell cell, Object value) {

        this.cell = cell;
        this.cellType = CstWorkCellType.VALUE;

        if (value == null) {
            return;
        }
        this.updateValue(value);
    }

    /**
     * Instantiates a new To work cell.
     *
     * @param cell  the cell
     * @param value the value
     * @param style the style
     */
    public CstWorkCell(@NonNull Cell cell, Object value, @NonNull CstWorkBookStyle style) {

        this(cell, value);
        this.cellType = CstWorkCellType.CUSTOM;
    }

    /**
     * Instantiates a new To work cell.
     *
     * @param cell  the cell
     * @param value the value
     * @param type  the type
     */
    public CstWorkCell(@NonNull Cell cell, Object value, @NonNull CstWorkCellType type) {
        this(cell, value);
        this.cellType = type;

    }



    private CstWorkCell updateValue(Object value) {
        if (value instanceof Double) {
            this.cell.setCellValue((double) value);
        } else if ( value instanceof Integer ){
            this.cell.setCellValue((int) value);
        } else if ( value instanceof Long ){
            this.cell.setCellValue((long) value);
        } else if (value instanceof Boolean) {
            this.cell.setCellValue((Boolean) value);
        } else if (value instanceof RichTextString) {
            this.cell.setCellValue((RichTextString) value);
        } else if (value instanceof Date) {
            this.cell.setCellValue((Date) value);
        } else if (value instanceof Calendar) {
            this.cell.setCellValue((Calendar) value);
        } else {
            this.cell.setCellValue(String.valueOf(value));
        }
        return this;
    }}
