package ru.nimdator;

import lombok.AllArgsConstructor;
import lombok.Getter;

@Getter
@AllArgsConstructor
public class ExcelProcessingException extends Exception {
    private Kind kind;
    private Integer row;
    private Integer column;

    @Getter
    @AllArgsConstructor
    public enum Kind {
        UNKNOWN("Неизвестная ошибка", -1),
        DATE_FORMAT_MISSMATCH("Неверный формат даты", -2),
        ;

        private final String description;
        private final Integer code;
    }

    @Override
    public String getMessage() {
        return kind.getDescription()
                + " Строка № " + (row == null ? "N/A" : row + 1)
                + " Колонка № " + (column == null ? "N/A" : column + 1);
    }

}
