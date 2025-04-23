package vn.com.phat.excel.param;

import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import java.util.List;
import java.util.function.Consumer;

/**
 * @author phatlt
 */
public interface ExcelExporterParam{
    List<ExcelExporterSheetParam> getSheets();
    String getDatePattern();
    Consumer<SXSSFWorkbook> getConsumer();
    String getOutputName();
    String getTemplatePath();

}
