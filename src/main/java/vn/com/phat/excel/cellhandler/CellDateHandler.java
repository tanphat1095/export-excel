package vn.com.phat.excel.cellhandler;

import org.apache.poi.xssf.streaming.SXSSFCell;

import java.util.Date;

public class CellDateHandler implements CellDataTypeHandler{

    @Override
    public void setCellData(SXSSFCell cell, Object value) {
        cell.setCellValue((Date) value);

    }
}
