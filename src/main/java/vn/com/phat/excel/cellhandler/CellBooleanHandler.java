package vn.com.phat.excel.cellhandler;

import org.apache.poi.xssf.streaming.SXSSFCell;

public class CellBooleanHandler implements CellDataTypeHandler{

    @Override
    public void setCellData(SXSSFCell cell, Object value) {
        cell.setCellValue(((Boolean) value).toString());
    }
}
