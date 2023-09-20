package ru.nimdator;

import lombok.extern.slf4j.Slf4j;

import org.apache.commons.math3.util.Pair;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;

import java.text.ParseException;
import java.text.SimpleDateFormat;

import java.util.Date;
import java.util.LinkedHashMap;
import java.util.LinkedList;
import java.util.List;
import java.util.Map;

import static ru.nimdator.ExcelConstants.EXCEL_DATE_FORMAT;

@Slf4j
public class ExcelMapper {


    private final Workbook workbook;
    private Sheet sheet;

    public ExcelMapper(String excelPath) throws ExcelFileMappingException {
        try (FileInputStream fileInputStream = new FileInputStream(excelPath)) {

            File file = new File(excelPath);
            if (!file.exists()) {
                log.info("File not found!!!");
                throw new FileNotFoundException(excelPath);
            }

            this.workbook = WorkbookFactory.create(fileInputStream);

            this.sheet = this.workbook.getSheetAt(0);
            if (this.sheet == null) {
                log.error("Sheet not found");
                throw new ExcelFileMappingException(ExcelFileMappingException.Kind.SHEET_NOT_FOUND);
            }
        } catch (FileNotFoundException e) {
            log.error("File not found at {}", excelPath);
            throw new ExcelFileMappingException(ExcelFileMappingException.Kind.FILE_NOT_FOUND, excelPath);
        } catch (SecurityException e) {
            log.error("SecurityException ", e);
            throw e;
        } catch (IOException e) {
            log.error("IOException ", e);
            throw new ExcelFileMappingException(ExcelFileMappingException.Kind.UNKNOWN);
        }
    }
    public <T> List<T> getFileContent (Class<T> cls) throws ExcelFileMappingException {
        String sheetName = ExcelObjectBuilder.getExcelSheetClass(cls);
        if (sheetName == null) {
            throw new ExcelFileMappingException(ExcelFileMappingException.Kind.CLASS_NOT_MARK_EXCEL);
        }
        setSheet(sheetName);
        Map<Field, Integer> mapHeader = ExcelObjectBuilder.getHeaderMapping(cls, getHeader());
        List<T> fileContent = new LinkedList<>();
        try {
            for (Row row : sheet) {
                if (row.getRowNum() == 0) {
                    continue;
                }
                List<Pair<Class<?>, String>> mapFieldValue = mapperClassFile(mapHeader, row);
                T obj = ExcelObjectBuilder.getObjFromMap(cls, mapFieldValue);
                fileContent.add(obj);
            }
        } catch (ParseException e) {
            log.error("ParseException ", e);
            throw new ExcelFileMappingException(ExcelFileMappingException.Kind.PARSING_EXEPTION);
        } catch (InvocationTargetException e) {
            log.error("InvocationTargetException ", e);
            throw new ExcelFileMappingException(ExcelFileMappingException.Kind.INVOCATION_EXEPTION);
        } catch (InstantiationException e) {
            log.error("InstantiationException ", e);
            throw new ExcelFileMappingException(ExcelFileMappingException.Kind.INSTANT_EXEPTION);
        } catch (IllegalAccessException e) {
            log.error("IllegalAccessException ", e);
            throw new ExcelFileMappingException(ExcelFileMappingException.Kind.ILLEGAL_ACCESS_EXEPTION);
        } catch (NoSuchMethodException e) {
            log.error("NoSuchMethodException ", e);
            throw new ExcelFileMappingException(ExcelFileMappingException.Kind.CONSTRUCTOR_NOT_FOUND);
        }
        return fileContent;
    }
    private <T> String getObjCellData(Cell cell, Class<T> requiredClass) {
        String cellData;
        log.info("rowno {} colno {} requiredClass {}", cell.getRowIndex(), cell.getColumnIndex(), requiredClass.getName());
        if (requiredClass.isAssignableFrom(Date.class)) {
            SimpleDateFormat formatter = new SimpleDateFormat(EXCEL_DATE_FORMAT);
            Date date = cell.getDateCellValue();
            cellData = formatter.format(date);
            log.info("DATE {} cell {}", date, cellData);
            return cellData;
        }
        if (requiredClass.isAssignableFrom(Integer.class)) {
            double value = cell.getNumericCellValue();
            cellData = Integer.toString((int) value);
            log.info("DATA {} cell {}", value, cellData);
            return cellData;
        }
        if (requiredClass.isAssignableFrom(Double.class)) {
            double value = cell.getNumericCellValue();
            cellData = Double.toString(value);
            log.info("DATA {} cell {}", value, cellData);
            return cellData;
        }

        return cell.getStringCellValue();
    }

    private List<Pair<String, Integer>> getHeader() {
        List<Pair<String, Integer>> headerList = new LinkedList<>();
        for (Cell cell : sheet.getRow(0)) {
            headerList.add(new Pair<>(cell.getStringCellValue(), cell.getColumnIndex()));
        }
        return headerList;
    }

    private void setSheet(String sheetName) {
        if (sheetName == null || sheetName.isEmpty()) {
            this.sheet = workbook.getSheetAt(0);
            return;
        }
        this.sheet = workbook.getSheet(sheetName);
    }
    private List<Pair<Class<?>, String>> mapperClassFile(Map<Field, Integer> mapHeader, Row row) {
        List<Pair<Class<?>, String>> mapFieldValue = new LinkedList<>();
        mapHeader.forEach((key, value) -> mapFieldValue.add(new Pair<>(key.getType(), getObjCellData(row.getCell(value), key.getType()))));
        return mapFieldValue;
    }
}
