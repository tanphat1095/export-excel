package vn.com.phat.excel;

import vn.com.phat.excel.exception.ExcelException;
import vn.com.phat.excel.param.ExcelExporterParam;

import java.io.OutputStream;

/**
 * Utility class for extracting data from objects for Excel processing.
 * This class provides methods to extract data from a list of objects, a single object, and to extract fields from a class.
 * The extracted data is used for processing Excel files.
 *
 * @author phatlt
 */
public interface ExcelExporter{
    /**
     * Performs the export of data to an Excel file.
     *
     * @param outStream The OutputStream where the Excel file will be written. This is used to write the Excel file.
     * @param param The parameters for the export. This includes the data to be exported and any additional settings.
     * @throws ExcelException If an error occurs during the export.
     */
    void doExport(OutputStream outStream, ExcelExporterParam param) throws ExcelException;
}
