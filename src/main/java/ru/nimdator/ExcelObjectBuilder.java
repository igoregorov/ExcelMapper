package ru.nimdator;

import lombok.extern.slf4j.Slf4j;
import org.apache.commons.math3.util.Pair;

import java.lang.reflect.Constructor;
import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

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
    public static <T> T getObjFromMap(Class<T> clz, List<Pair<Class<?>, String>> classStringMap)
            throws InvocationTargetException, InstantiationException, IllegalAccessException, ParseException, NoSuchMethodException {

        Object[] objects = new Object[classStringMap.size()];
        Class<?>[] classes = new Class[classStringMap.size()];
        int i = 0;
        for (Pair<Class<?>, String> entry : classStringMap) {
            Class<?> fieldClass = entry.getFirst();
            objects[i] = getObjFromString(fieldClass, entry.getSecond());
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

    public static Map<Field, Integer> getHeaderMapping(Class<?> cls, List<Pair<String, Integer>> headerList) throws ExcelFileMappingException {
        Map<Field, Integer> mapFields = new LinkedHashMap<>();
        for (Field field : cls.getDeclaredFields()) {
            ExcelColumn excelColumn = field.getAnnotation(ExcelColumn.class);
            if (excelColumn == null) {
                log.info("column is null for field {}", field.getName());
                continue;
            }
            for (Pair<String, Integer> cell : headerList) {
                if (cell.getFirst().equals(excelColumn.name()) || cell.getSecond() == excelColumn.index()) {
                    mapFields.put(field, cell.getSecond());
                    break;
                }
            }
            if (!mapFields.containsKey(field) && !cls.getAnnotation(ExcelSheet.class).ignoreUnkonow()) {
                throw new ExcelFileMappingException(ExcelFileMappingException.Kind.COLUMN_NOT_EXISTS, excelColumn.name());
            }
        }
        return mapFields;
    }
}
