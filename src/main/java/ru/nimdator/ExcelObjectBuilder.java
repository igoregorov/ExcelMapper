package ru.nimdator;

import lombok.extern.slf4j.Slf4j;


import java.lang.reflect.Constructor;
import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.LinkedList;
import java.util.List;

import static ru.nimdator.ExcelConstants.EXCEL_DATE_FORMAT;

@Slf4j
public class ExcelObjectBuilder {
    private ExcelObjectBuilder() {
        throw new IllegalStateException("Utility class");
    }

    /**
     * Метод возвращает экземпляр заданного класса.
     * Метод находит конструктор, содержащий типы, перечисленные в entryKey
     * classStringMap.
     * Важно, чтобы при заполнении classStringMap, сохранялся порядок типов
     * порождающего класса. Рекомендуется использовать LinkedHashMap.
     * @param clz - порождаюбщий класс
     * @param classStringMap - именованный упорядоченный словарь полеКласса-Значение
     * @throws InvocationTargetException
     * @throws InstantiationException
     * @throws IllegalAccessException
     * @throws ParseException
     * @throws NoSuchMethodException
     */
    public static <T> T getObjFromMap(Class<T> clz, List<ExcelPair<Class<?>, String>> classStringMap)
            throws InvocationTargetException, InstantiationException, IllegalAccessException, ParseException, NoSuchMethodException {

        Object[] objects = new Object[classStringMap.size()];
        Class<?>[] classes = new Class[classStringMap.size()];
        int i = 0;
        for (ExcelPair<Class<?>, String> entry : classStringMap) {
            Class<?> fieldClass = entry.getKey();
            objects[i] = getObjFromString(fieldClass, entry.getValue());
            classes[i++] = fieldClass;
        }
        return clz.cast(clz.getDeclaredConstructor(classes).newInstance(objects));
    }

    public static <T> T getObjFromString(Class<T> clz, String obj) throws InvocationTargetException, InstantiationException, IllegalAccessException, ParseException {
        log.info("Operation at {}", obj);
        Constructor<?> stringConstructor = null;
        if (clz.isAssignableFrom(String.class)) {
            return clz.cast(obj);
        }
        if (clz.isAssignableFrom(Date.class)) {
            SimpleDateFormat formatter = new SimpleDateFormat(EXCEL_DATE_FORMAT);
            return clz.cast(formatter.parse(obj));
        }
        for (Constructor<?> constructor : clz.getDeclaredConstructors()) {
            Class<?>[] clses = constructor.getParameterTypes();
            if (clses.length == 1 && String.class.isAssignableFrom(clses[0])) {
                stringConstructor = constructor;
            }
        }
        if (stringConstructor == null) {
            throw new IllegalAccessException("Класс не имеет конструктора String");
        }
        return clz.cast(stringConstructor.newInstance(obj));
    }

    public static String getExcelSheetClass(Class<?> cls) {
        ExcelSheet sheetName = cls.getAnnotation(ExcelSheet.class);
        if (sheetName == null) {
            return null;
        }
        return sheetName.name();
    }

    public static List<ExcelPair<Class<?>, Integer>> getHeaderMapping(Class<?> cls, List<ExcelPair<String, Integer>> headerList) throws ExcelFileMappingException {
        List<ExcelPair<Class<?>, Integer>> mapFields = new LinkedList<>();
        for (Field field : cls.getDeclaredFields()) {
            ExcelColumn excelColumn = field.getAnnotation(ExcelColumn.class);
            boolean isFound = false;
            if (excelColumn == null) {
                log.info("column is null for field {}", field.getName());
                continue;
            }
            for (ExcelPair<String, Integer> cell : headerList) {
                if (cell.getKey().equals(excelColumn.name()) || cell.getValue() == excelColumn.index()) {
                    mapFields.add(new ExcelPair<>(field.getType(), cell.getValue()));
                    isFound = true;
                    break;
                }
            }
            if (!isFound && !cls.getAnnotation(ExcelSheet.class).ignoreUnkonow()) {
                throw new ExcelFileMappingException(ExcelFileMappingException.Kind.COLUMN_NOT_EXISTS, excelColumn.name());
            }
        }
        return mapFields;
    }
}
