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
package vn.com.phat.excel.util;

import vn.com.phat.excel.dto.ItemColsExcelDto;

import java.util.List;

/**
 * @author phatlt
 */
public class ExcelColumnExtractor {

    private ExcelColumnExtractor(){}

    public static <E extends Enum<E> & EnumExcel> void extractColumn(Class<E> enumType, List<ItemColsExcelDto> cols) {
        int index = -1;
        for (E en : enumType.getEnumConstants()) {
            index = en.getIndex() >= 0 ? en.getIndex() : index + 1;
            ItemColsExcelDto col = new ItemColsExcelDto(index, en.name());
            cols.add(col);
        }

    }
}
