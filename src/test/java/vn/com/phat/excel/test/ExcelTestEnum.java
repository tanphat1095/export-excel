package vn.com.phat.excel.test;

import vn.com.phat.excel.util.EnumExcel;

public enum ExcelTestEnum implements EnumExcel {
    NO(-1),
    NAME(-1),
    DATE(-1),
    ISTRUE(-1),

    LOCALDATE(10)

    , AMOUNT(-1)
    , CUSTOMSTYLE(-1)
    ;

    ExcelTestEnum(Integer index){
        this.index = index;
    }

    private final Integer index;

    @Override
    public Integer getIndex() {
        return this.index;
    }
}
