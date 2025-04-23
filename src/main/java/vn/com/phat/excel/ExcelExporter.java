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
 * Exports data to an Excel file and writes the result to the provided output stream.
 *
 * @param outStream the output stream to which the Excel file content will be written
 * @param param encapsulates the data and configuration for the export operation
 * @throws ExcelException if an error occurs during the export process
 */
    void doExport(OutputStream outStream, ExcelExporterParam param) throws ExcelException;
}
