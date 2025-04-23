package vn.com.phat.excel.cellhandler;

import org.apache.poi.xssf.streaming.SXSSFCell;

public class CellStringHandler implements CellDataTypeHandler {

    @Override
    public void setCellData(SXSSFCell cell, Object value) {
        cell.setCellValue((String) value);
    }
}
