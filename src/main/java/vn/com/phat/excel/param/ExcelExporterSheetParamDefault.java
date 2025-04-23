package vn.com.phat.excel.param;

import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Getter;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import vn.com.phat.excel.dto.ItemColsExcelDto;

import java.lang.reflect.Field;
import java.util.List;
import java.util.Map;
import java.util.function.Consumer;

/**
 * @author phatlt
 */
@Getter
@AllArgsConstructor
@Builder
public class ExcelExporterSheetParamDefault implements ExcelExporterSheetParam{

    private final List<Map<String, Object>> data;
    private final int sheetIndex;
    private final List<ItemColsExcelDto> colIndex;
    private final Map<String, Field> mapFields;
    private final String startCell;
    private final Consumer<SXSSFSheet> consumer;
    private final String sheetName;
}
