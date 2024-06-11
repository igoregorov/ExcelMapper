package ru.nimdator;

import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.util.ArrayList;
import java.util.Date;
import java.util.LinkedList;
import java.util.List;

@Slf4j
public class ExcelMapper {
    private final Workbook workbook;
    public ExcelMapper(String excelPath) throws ExcelFileMappingException {
        workbook = new ExcelFile(excelPath).getWorkbook();
    }
    public <T> List<T> getFileContent(Class<T> cls) throws ExcelFileMappingException {

        ObjectFields objectFields = new ObjectFields(cls);
        List<ColumnProperty> propertyList  = objectFields.getColumnProperties();
        SheetProperty sheetProperty = objectFields.getSheetProperty();

        List<Integer> mapHeader = new ArrayList<>();
        List<T> fileAsList = new LinkedList<>();

        for (Row row : getSheetByName(sheetProperty.getSheetName())) {
            if (row.getRowNum() == 0 && sheetProperty.isReadHeader()) {
                mapHeader = ExcelObjectBuilder.getHeaderMapping(getHeader(row), propertyList, sheetProperty.isIgnoreUnknown());
                continue;
            }
            List<Object> mapFieldValue = mapperClassFile(mapHeader, row, propertyList);
            T obj = ExcelObjectBuilder.getObjFromMap(cls, mapFieldValue, propertyList);
            fileAsList.add(obj);
        }

        return fileAsList;
    }

    private List<String> getHeader(Row row) throws ExcelFileMappingException {
        List<String> headerList = new LinkedList<>();
        for (Cell cell : row) {
            headerList.add(getCellValue(cell, String.class));
        }
        log.debug("header is {}", headerList);
        return headerList;
    }

    private Sheet getSheetByName(String sheetName) throws ExcelFileMappingException {
        if (sheetName == null || sheetName.isEmpty()) {
            return workbook.getSheetAt(0);
        }
        Sheet sheet =workbook.getSheet(sheetName);
        if (sheet == null) {
            throw new ExcelFileMappingException(ExcelFileMappingException.Kind.SHEET_NOT_FOUND, sheetName);
        }
        return sheet;
    }

    private List<Object> mapperClassFile(List<Integer> mapHeader, Row row, List<ColumnProperty> propertyList) throws ExcelFileMappingException {
        List<Object> mapFieldValue = new LinkedList<>();
        for (int i = 0; i < mapHeader.size(); i++) {
            Integer colPosition = mapHeader.get(i);
            if (colPosition == -1) {
                mapFieldValue.add(null);
                continue;
            }
            Cell cell = row.getCell(colPosition);
            Class<?> reqClass = propertyList.get(i).getTClass();
            log.debug("Cell col {} and class is {}", (cell != null ? cell.getColumnIndex() : null), (cell != null ? reqClass.getName() : null));
            log.debug("Cell value is {}", getCellValue(cell, reqClass));
            mapFieldValue.add(getCellValue(cell, reqClass));
        }
        return mapFieldValue;
    }
    private <T> T getCellValue(Cell cell, Class<T> reqClass) throws ExcelFileMappingException {
        if (cell == null) {
            return null;
        }
        try {
            return getObjCellData(cell, reqClass);
        } catch (IllegalStateException | NumberFormatException | ClassCastException e) {
            log.error("Parsing Error rowno {} colno {} requiredClass {}", cell.getRowIndex(), cell.getColumnIndex(), reqClass.getName());
            log.error("IllegalStateException {}", e.getMessage());
            log.error("", e);
            throw new ExcelFileMappingException(ExcelFileMappingException.Kind.ERROR_PARSE_CELL,
                    e.getMessage(),
                    String.valueOf(cell.getRowIndex()),
                    String.valueOf(cell.getColumnIndex()),
                    reqClass.getName());
        }
    }
    private <T> T getObjCellData(Cell cell, Class<T> requiredClass) {
        if (cell == null) {
            return null;
        }
        if (requiredClass.isAssignableFrom(Date.class)) {
            return CellData.getObjCellData(cell.getDateCellValue(), requiredClass);
        }

        Object obj;

        switch (cell.getCellType()) {
            case _NONE, FORMULA, BOOLEAN, ERROR -> log.debug("null");
            case NUMERIC -> {
                obj = CellData.getObjCellData(cell.getNumericCellValue(), requiredClass);
                log.debug("NUMERIC {}", obj);

            }
            case BLANK -> {
                obj = CellData.getObjCellData(requiredClass);
                log.debug("BLANK {}", obj);
            }
            default -> {
                obj = CellData.getObjCellData(cell.getStringCellValue(), requiredClass);
                log.debug("DEFAULT {}", obj);
            }
        }

        return requiredClass.cast(
                switch (cell.getCellType()) {
                    case _NONE, FORMULA, BOOLEAN, ERROR -> null;
                    case NUMERIC -> CellData.getObjCellData(cell.getNumericCellValue(), requiredClass);
                    case BLANK -> CellData.getObjCellData(requiredClass);
                    default -> CellData.getObjCellData(cell.getStringCellValue(), requiredClass);
                });
    }
}
