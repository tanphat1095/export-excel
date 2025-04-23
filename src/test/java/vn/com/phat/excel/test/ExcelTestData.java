package vn.com.phat.excel.test;

import lombok.AllArgsConstructor;
import lombok.Getter;

import java.math.BigDecimal;
import java.time.LocalDate;
import java.util.Date;

@Getter
@AllArgsConstructor
public class ExcelTestData {
    private final Integer no;
    private final String name;
    private final Date date;
    private final Boolean isTrue;
    private final LocalDate localDate;
    private final BigDecimal amount;
    private final String customStyle;
}
