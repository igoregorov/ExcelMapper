package ru.nimdator;

import lombok.Getter;
import lombok.RequiredArgsConstructor;
import lombok.ToString;

import java.util.Date;

@ExcelSheet(ignoreUnkonow = true)
@Getter
@RequiredArgsConstructor
@ToString
public class ExcelTestFile {
    @ExcelColumn(name = "test2")
    private final Integer fld2;
    @ExcelColumn(name = "test3")
    private final Date fld3;
    @ExcelColumn(name = "test1")
    private final String fld1;
    @ExcelColumn(name = "test4")
    private final Double fld4;
    @ExcelColumn(name = "test5")
    private final Double fld5;
    @ExcelColumn(name = "tes6")
    private final String fld6;
    @ExcelColumn(name = "ИНН")
    private final String colConsignorInn;

}
