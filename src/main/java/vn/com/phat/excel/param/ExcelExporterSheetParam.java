package vn.com.phat.excel.param;

import org.apache.poi.xssf.streaming.SXSSFSheet;
import vn.com.phat.excel.dto.ItemColsExcelDto;

import java.lang.reflect.Field;
import java.util.List;
import java.util.Map;
import java.util.function.Consumer;

/**
 * @author phatlt 
 */
public interface ExcelExporterSheetParam {
    List<Map<String, Object>> getData();
    int getSheetIndex();
    List<ItemColsExcelDto> getColIndex();
    Map<String, Field> getMapFields();
    String getStartCell();
    Consumer<SXSSFSheet> getConsumer();
    String getSheetName();
}
