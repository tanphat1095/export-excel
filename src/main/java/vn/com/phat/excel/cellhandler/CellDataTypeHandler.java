package vn.com.phat.excel.cellhandler;

import org.apache.poi.xssf.streaming.SXSSFCell;


/**
 * This interface defines a method for handling cell data types in an Excel workbook.
 * Implementations of this interface should provide a specific way to set data for a cell.
 *
 * @author phatlt
 */
public interface CellDataTypeHandler{
    void setCellData(SXSSFCell cell, Object value);
}
