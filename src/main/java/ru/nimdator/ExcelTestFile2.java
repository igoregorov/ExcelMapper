package ru.nimdator;

import lombok.AllArgsConstructor;
import lombok.Getter;
import lombok.Setter;
import lombok.ToString;

@ExcelSheet
@Getter
@Setter
@AllArgsConstructor
@ToString
public class ExcelTestFile2 {
    @ExcelColumn(name = "test2")
    private Integer fld2;
/*
    public ExcelTestFile(String fld1, Integer fld2) {
        this.fld1 = fld1;
        this.fld2 = fld2;
    }

 */
}
