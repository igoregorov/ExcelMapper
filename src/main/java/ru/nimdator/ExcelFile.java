package ru.nimdator;

import lombok.NoArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;


import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Locale;

import static org.apache.poi.ss.usermodel.CellType.STRING;

@NoArgsConstructor
@Slf4j
public class ExcelFile {

    private static final String EXCEL_DATE_FORMAT = "dd.MM.yyyy";

    private Workbook workbook;
    private Sheet sheet;

    public void info() {
        log.info("Egorov haha");
    }

    public String getCellData(int rowno, int colno) {
        try {
            Cell cell = sheet.getRow(rowno).getCell(colno);
            String CellData = null;
            switch (cell.getCellType()) {
                case STRING:
                    CellData = cell.getStringCellValue();
                    break;
                case NUMERIC:
                    // Из ячейки считывается код завода 2175.0, а должно 2175
                    Double dataDouble = cell.getNumericCellValue();
                    CellData = Integer.toString(dataDouble.intValue());
                    break;
                case BLANK:
                    CellData = "";
                    break;
            }
            return CellData;
        } catch (Exception e) {
            return "";
        }
    }

    public String getCellValueAsString(int rowNumber, int columnNumber) throws ExcelProcessingException {
        try {
            Cell cell = sheet.getRow(rowNumber).getCell(columnNumber);
            return new DataFormatter().formatCellValue(cell);
        } catch (Exception e) {
            log.error("getCellValueAsString() error", e);
            throw new ExcelProcessingException(ExcelProcessingException.Kind.UNKNOWN, rowNumber, columnNumber);
        }
    }

    public Date getCellValueAsDate(int rowNumber, int columnNumber) {
        Cell cell = sheet.getRow(rowNumber).getCell(columnNumber);
        String cellValue = new DataFormatter(new Locale("ru")).formatCellValue(cell);
        if (cellValue == null || cellValue.isEmpty()) {
            return null;
        }
        if (DateUtil.isCellDateFormatted(cell)) {
            log.debug("CellDateFormatted cellValue = {}", new SimpleDateFormat(EXCEL_DATE_FORMAT).format(cell.getDateCellValue()));
            return cell.getDateCellValue();
        } else {
            throw new IllegalArgumentException("Excel cell [" + rowNumber + "]" + "[" + columnNumber + "]" + "must be date formatted");
        }
    }

    public Double getCellValueAsDouble(int rowNumber, int columnNumber) throws ExcelProcessingException {
        Cell cell = sheet.getRow(rowNumber).getCell(columnNumber);
        String cellValue = (new DataFormatter().formatCellValue(cell)).replace(",", ".");
        try {
            return Double.valueOf(cellValue);
        } catch (NumberFormatException e) {
            log.error("getCellValueAsDouble() error", e);
            throw new ExcelProcessingException(ExcelProcessingException.Kind.UNKNOWN, rowNumber, columnNumber);
        }
    }

    public void setNewExcelFile(String excelFile) {
        try {
            File file = new File(excelFile);
            workbook = WorkbookFactory.create(file);

        } catch (IOException | EncryptedDocumentException e) {
            log.error("error", e);
        }
    }

    public void setExcelFile(String excelPath) throws Exception {
        try (FileInputStream fileInputStream = new FileInputStream(excelPath)) {

            File file = new File(excelPath);
            if (!file.exists()) {
                log.info("File not found!!!");
                throw new FileNotFoundException(excelPath);
            }

            workbook = WorkbookFactory.create(fileInputStream);
            if (workbook == null) {
                log.info("Workbook not found");
                throw new RuntimeException("WorkBook not found");
            }

            sheet = workbook.getSheetAt(0);
            if (sheet == null) {
                log.info("Sheet not found");
                throw new RuntimeException("Sheet not found");
            }
        } catch (Exception e) {
            log.error("Error pasing excel file: ", e);
            throw e;
        }
    }

    public int getLastRow() {
        return sheet.getLastRowNum();
    }

    public boolean setSheetByName(String sheetName) {
        if (workbook == null) {
            log.info("Workbook not found");
            throw new RuntimeException("WorkBook not found");
        }
        sheet = workbook.getSheet(sheetName);
        return sheet != null;
    }

    public int getColumnByHeaderName(String headerName) {
        Row row = sheet.getRow(0);
        for (Cell cell1 : row) {
            if (cell1.getCellType() == STRING && headerName.equals(cell1.getStringCellValue())) {
                return cell1.getColumnIndex();
            }
        }
        return -1;
    }
}
