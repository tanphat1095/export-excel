package vn.com.phat.excel.cellhandler;

import org.apache.poi.xssf.streaming.SXSSFCell;

public class CellIntegerHandler implements CellDataTypeHandler{

    @Override
    public void setCellData(SXSSFCell cell, Object value) {
        cell.setCellValue(Integer.parseInt(value.toString()));
    }
}
