package ru.nimdator;

import lombok.Getter;
import lombok.RequiredArgsConstructor;
import org.apache.poi.ss.usermodel.Sheet;

@Getter
@RequiredArgsConstructor
public class SheetProperty {

    private final String sheetName;
    private final boolean readHeader;
    private final boolean ignoreUnknown;

}
