package vn.com.phat.excel.cellstyle;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

/**
 * @author phatlt
 */
public class CellStyleBigDecimalHandler implements CellStyleHandler{

    @Override
    public CellStyle handleCellStyle(SXSSFWorkbook workbook, DataFormat dataFormat) {
        CellStyle cellStyle = workbook.createCellStyle();
        cellStyle.setAlignment(HorizontalAlignment.RIGHT);
        cellStyle.setDataFormat(workbook.createDataFormat().getFormat("#,##0.00"));
        return cellStyle;
    }

}
