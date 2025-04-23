package vn.com.phat.excel.cellhandler;

import org.apache.poi.xssf.streaming.SXSSFCell;

import java.math.BigDecimal;
import java.math.RoundingMode;

public class CellBigDecimalHandler implements CellDataTypeHandler{

    @Override
    public void setCellData(SXSSFCell cell, Object value) {
        cell.setCellValue(((BigDecimal) value).setScale(2, RoundingMode.UP).doubleValue());
    }
}
