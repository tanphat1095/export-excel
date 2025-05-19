/*
 * Copyright 2025 tanphat.1095
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */
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
