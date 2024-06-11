package ru.nimdator;

import lombok.Getter;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

@Slf4j
public final class ExcelFile {
    @Getter
    private final Workbook workbook;

    public ExcelFile(String excelPath) throws ExcelFileMappingException {
        try (FileInputStream fileInputStream = new FileInputStream(excelPath)) {

            File file = new File(excelPath);
            if (!file.exists()) {
                log.info("File not found!!!");
                throw new FileNotFoundException(excelPath);
            }

            this.workbook = WorkbookFactory.create(fileInputStream);

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
}
