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
package vn.com.phat.excel.cellstyle;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

/**
 * @author phatlt
 * This interface defines a method for handling cell styles in an Excel workbook.
 * Implementations of this interface should provide a specific way to create and configure a CellStyle object.
 */


public interface CellStyleHandler {

    /**
     * Handles the creation and configuration of a CellStyle object.
     *
     * @param workbook The workbook where the cell style will be applied. This is used to create the CellStyle object.
     * @param dataFormat The data format to be applied to the cell style. This is used to set the data format of the CellStyle object.
     * @return A CellStyle object configured with the provided data format.
     */
    CellStyle handleCellStyle(SXSSFWorkbook workbook, DataFormat dataFormat);

}
