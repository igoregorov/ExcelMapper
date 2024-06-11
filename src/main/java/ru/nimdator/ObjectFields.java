package ru.nimdator;

import lombok.Getter;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.lang.reflect.Field;
import java.lang.reflect.Modifier;
import java.util.ArrayList;
import java.util.LinkedList;
import java.util.List;

import static ru.nimdator.ExcelFileMappingException.Kind.EXCEL_COLUMN_OR_FINAL;

@Slf4j
public final class  ObjectFields {
    @Getter
    private final List<ColumnProperty> columnProperties;
    @Getter
    private final SheetProperty sheetProperty;
    public ObjectFields(Class<?> tClass) throws ExcelFileMappingException {
        this.columnProperties = getAnnotationValues(tClass);
        this.sheetProperty = getSheetByName(tClass);
    }

    private List<ColumnProperty> getAnnotationValues(Class<?> tClass) throws ExcelFileMappingException {
        List<ColumnProperty> propertyList = new LinkedList<>();
        for (Field field : tClass.getDeclaredFields()) {
            boolean isFinal = Modifier.isFinal(field.getModifiers());
            boolean isExcelColumn = field.isAnnotationPresent(ExcelColumn.class);
            if (!isExcelColumn && !isFinal) continue;
            if (!isExcelColumn || !isFinal) throw new ExcelFileMappingException(EXCEL_COLUMN_OR_FINAL, field.getName());
            ExcelColumn column = field.getAnnotation(ExcelColumn.class);
            if ((column.name() == null || column.name().isEmpty()) && column.index() == -1) continue;
            ColumnProperty columnProperty = new ColumnProperty(column.name(), column.index(), field.getType());
            propertyList.add(columnProperty);
        }
        return propertyList;
    }

    private SheetProperty getSheetByName(Class<?> cls) throws ExcelFileMappingException {
        ExcelSheet excelSheet = cls.getAnnotation(ExcelSheet.class);
        if (excelSheet == null) {
            throw new ExcelFileMappingException(ExcelFileMappingException.Kind.CLASS_NOT_MARK_EXCEL);
        }
        String sheet = excelSheet.name();
        boolean ignoreUnknown = cls.getAnnotation(ExcelSheet.class).ignoreUnkonow();
        boolean readHeader = cls.getAnnotation(ExcelSheet.class).readHeader();
        return new SheetProperty(sheet, readHeader, ignoreUnknown);
    }

}
