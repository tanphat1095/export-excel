package vn.com.phat.excel.param;

import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Getter;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import java.util.List;
import java.util.function.Consumer;

/**
 * @author phatlt
 */
@Getter
@AllArgsConstructor
@Builder
public class ExcelExporterParamDefault implements ExcelExporterParam{

    private final List<ExcelExporterSheetParam> sheets;
    private final String datePattern;
    private final Consumer<SXSSFWorkbook> consumer;
    private final String outputName;
    private final String templatePath;
}
