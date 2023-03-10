package com.excel.core.type;

/**
 * The enum To work cell type.
 */
public enum CstWorkCellType {

    /**
     * Title to work cell type.
     */
    TITLE,
    /**
     * Value to work cell type.
     */
    VALUE,
    /**
     * Custom to work cell type.
     */
    CUSTOM;


    /**
     * Is title boolean.
     *
     * @return the boolean
     */
    public boolean isTitle(){
        return TITLE.equals(this);
    }
}
