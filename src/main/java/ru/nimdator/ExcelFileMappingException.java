package ru.nimdator;

import lombok.AllArgsConstructor;
import lombok.Getter;

import java.util.IllegalFormatException;

@Getter
public class ExcelFileMappingException extends Exception {
    private final Kind kind;

    public ExcelFileMappingException(Kind kind) {
        super(kind.description);
        this.kind = kind;
    }
    public ExcelFileMappingException(Kind kind, String... info) {
        super(getFormattedCode(kind.getDescription(), info));
        this.kind = kind;
    }

    @Getter
    @AllArgsConstructor
    public enum Kind {
        UNKNOWN("Неизвестная ошибка", -1),
        DATE_FORMAT_MISSMATCH("Неверный формат даты", -2),
        SHEET_NOT_FOUND("Не найден лист %s в книге Экселя", -3),
        FILE_NOT_FOUND("Не найден файл %s", -4),
        PARSING_EXEPTION("Ошибка при парсинге файла", -5),
        INVOCATION_EXEPTION("Недопустимая операция при вызове конструктора", -6),
        INSTANT_EXEPTION("Объект не может быть создан", -7),
        ILLEGAL_ACCESS_EXEPTION("Попытка вызвать метод или конструктор без права доступа", -8),
        WORKBOOOK_NOT_FOUND("Workbook не найден", -9),
        CONSTRUCTOR_NOT_FOUND("Не задан конструктор класса", -10),
        CLASS_NOT_MARK_EXCEL("Класс не помечен аннотацией ExcelSheet",-11),
        COLUMN_NOT_EXISTS("В файле не найдена колонка %s",-12),
        ERROR_PARSE_CELL("Ошибка %s строка %s колонка %s ожидаемый тип данных %s", -13),
        EXCEL_COLUMN_OR_FINAL("Поле класса %s или final или не final но с аннотацией ExcelColumn", -14),
        ;

        private final String description;
        private final Integer code;
    }

    private static String getFormattedCode(String statusMessage, String[] placeholders) {
        if (statusMessage == null)
            return null;
        try {
            return String.format(statusMessage, (Object[]) placeholders);
        } catch (IllegalFormatException e) {
            return statusMessage;
        }
    }


}
