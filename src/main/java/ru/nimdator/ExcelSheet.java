package ru.nimdator;

import java.lang.annotation.Retention;
import java.lang.annotation.Target;

import static java.lang.annotation.ElementType.TYPE;
import static java.lang.annotation.RetentionPolicy.RUNTIME;

/**
 * Указывает, что класс описывает лист Excel файла.
 * <p> Пример класса с аннотацией {@code ExcelSheet}
 *
 * <blockquote><pre>{@code
 * @ExcelSheet
 * @Getter
 * @RequiredArgsConstructor
 * @ToString
 * public class ExcelTestFile {
 *
 *      @ExcelColumn(name = "test1")
 *      private final String fld1;
 *
 *      @ExcelColumn(name = "test2")
 *      private final Integer fld2;
 *
 *      @ExcelColumn(name = "test3")
 *      private final Date fld3;
 *
 *      private Double fld4;
 *}
 *}</pre></blockquote>
 * @since 90.0
 */
@Retention(RUNTIME)
@Target(TYPE)
public @interface ExcelSheet {
    /**
     * Не обязательный параметр названия конкретного листа файла
     */
    String name() default "";
    boolean ignoreUnkonow() default false;
}
