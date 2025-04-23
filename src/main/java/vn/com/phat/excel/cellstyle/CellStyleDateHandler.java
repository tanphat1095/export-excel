package vn.com.phat.excel.cellstyle;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.ss.usermodel.HorizontalAlignment;

/**
 * @author phatlt
 */
public class CellStyleDateHandler implements CellStyleHandler {

    @Override
    public CellStyle handleCellStyle(SXSSFWorkbook workbook, DataFormat dataFormat) {
        CellStyle cellStyle = workbook.createCellStyle();
        cellStyle.setAlignment(HorizontalAlignment.CENTER);
        cellStyle.setDataFormat(workbook.createDataFormat().getFormat("dd/MM/yyyy"));
        return cellStyle;
    }

}
