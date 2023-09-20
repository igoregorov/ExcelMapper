package ru.nimdator;

import lombok.extern.slf4j.Slf4j;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.usermodel.Sheet;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.List;

@Slf4j
public class NewExcelFile {
    private final Workbook workbook;
    public NewExcelFile(String filePath) throws IOException {
        InputStream inputStream = new FileInputStream(filePath);
        workbook = WorkbookFactory.create(inputStream);
    }
    public List<Sheet> getListSheets() {
        List<Sheet> sheets = new ArrayList<>();
        for (Sheet sheet : workbook) {
            sheets.add(sheet);
        }
        log.info("Sheets size = {}", sheets.size());
        return sheets;
    }

    public void mapperExcel() {
        Sheet sheet = workbook.getSheetAt(0);
        for (Row row : sheet) {

        }
    }

    public <T> void processRec(Class<T> tClass) {

    }

    public void searchTextExcel(Class<?> clazz) {
        for (Field field : clazz.getDeclaredFields()) {
            ExcelColumn excelColumn = field.getAnnotation(ExcelColumn.class);
            if (excelColumn == null) {
                continue;
            }
            log.info("field {} - {} - {}", field.getName(), excelColumn.name(), field.getType().getName());
        }
    }

    public static boolean set(Object object, String fieldName, Object fieldValue) {
        Class<?> clazz = object.getClass();
        while (clazz != null) {
            try {
                Field field = clazz.getDeclaredField(fieldName);

                return true;
            } catch (NoSuchFieldException e) {
                clazz = clazz.getSuperclass();
            } catch (Exception e) {
                throw new IllegalStateException(e);
            }
        }
        return false;
    }
}
