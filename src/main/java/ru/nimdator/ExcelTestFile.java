package ru.nimdator;

import lombok.AllArgsConstructor;
import lombok.Getter;
import lombok.NonNull;
import lombok.RequiredArgsConstructor;
import lombok.Setter;
import lombok.ToString;

import java.util.Date;

@ExcelSheet
@Getter
@RequiredArgsConstructor
@ToString
public class ExcelTestFile {
    @ExcelColumn(name = "test1")
    private final String fld1;
    @ExcelColumn(name = "test2")
    private final Integer fld2;
    @ExcelColumn(name = "test3")
    private final Date fld3;
    private Double fld4;
}
