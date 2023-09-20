package ru.nimdator;

import lombok.extern.slf4j.Slf4j;

import java.lang.reflect.Constructor;
import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Date;
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
    public static <T> T getObjFromMap(Class<T> clz, Map<Field, String> classStringMap) throws InvocationTargetException, InstantiationException, IllegalAccessException, ParseException, NoSuchMethodException {

        Object[] objects = new Object[classStringMap.size()];
        Class<?>[] classes = new Class[classStringMap.size()];
        int i = 0;
        for (Map.Entry<Field, String> entry : classStringMap.entrySet()) {
            Class<?> fieldClass = entry.getKey().getType();
            String value = entry.getValue();
            objects[i] = getObjFromString(fieldClass, value);
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
}
