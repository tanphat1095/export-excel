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
