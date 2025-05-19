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
package vn.com.phat.excel.cellhandler;

import org.apache.poi.xssf.streaming.SXSSFCell;


/**
 * This interface defines a method for handling cell data types in an Excel workbook.
 * Implementations of this interface should provide a specific way to set data for a cell.
 *
 * @author phatlt
 */
public interface CellDataTypeHandler{
    void setCellData(SXSSFCell cell, Object value);
}
