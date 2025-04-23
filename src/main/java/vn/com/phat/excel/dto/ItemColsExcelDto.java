package vn.com.phat.excel.dto;

import lombok.AllArgsConstructor;
import lombok.Getter;

@Getter
@AllArgsConstructor
public class ItemColsExcelDto {
    private final Integer colIndex;
    private final String colName;

    public ItemColsExcelDto(int colIndex, String colName){
        this.colIndex = colIndex;
        this.colName = colName;
    }
}
