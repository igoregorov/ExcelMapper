package ru.nimdator;

import lombok.extern.slf4j.Slf4j;

import java.lang.reflect.InvocationTargetException;
import java.text.ParseException;
import java.util.LinkedList;
import java.util.List;

@Slf4j
public final class ExcelObjectBuilder {
    private ExcelObjectBuilder() {
        throw new IllegalStateException("Utility class");
    }

    /**
     * Метод возвращает экземпляр заданного класса.
     * Метод находит конструктор, содержащий типы, перечисленные в entryKey
     * classStringMap.
     * Важно, чтобы при заполнении classStringMap, сохранялся порядок типов
     * порождающего класса. Рекомендуется использовать LinkedHashMap.
     * @param tClass - порождаюбщий класс
     * @param classStringMap - именованный упорядоченный словарь полеКласса-Значение
     * @throws InvocationTargetException
     * @throws InstantiationException
     * @throws IllegalAccessException
     * @throws ParseException
     * @throws NoSuchMethodException
     */
    public static <T> T getObjFromMap(Class<T> tClass, List<Object> classStringMap, List<ColumnProperty> columnProperty) throws ExcelFileMappingException {
        Object[] objects = new Object[classStringMap.size()];
        Class<?>[] classes = new Class[classStringMap.size()];
        int i = 0;
        try {
            for (Object entry : classStringMap) {
                objects[i] = entry;
                classes[i] = columnProperty.get(i++).getTClass();
            }
        return tClass.cast(tClass.getDeclaredConstructor(classes).newInstance(objects));
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
    }

    public static List<Integer> getHeaderMapping(List<String> headerList, List<ColumnProperty> propertyList, boolean ignoreUnknown) throws ExcelFileMappingException {
        List<Integer> mapFields = new LinkedList<>();
        for (ColumnProperty field : propertyList) {
            String name = field.getColumnName();
            int index = field.getColumnIndex();
            boolean isFound = false;
            for (int i =0; i < headerList.size(); i++) {
                String headerCol = headerList.get(i);
                if (headerCol.equals(name) || i == index) {
                    log.debug("Headet col {} has index {} mapped with field {} at class {}", headerCol, i, name, field.getTClass());
                    mapFields.add(i);
                    isFound = true;
                    break;
                }
            }
            if (isFound) continue;
            if (!ignoreUnknown) {
                throw new ExcelFileMappingException(ExcelFileMappingException.Kind.COLUMN_NOT_EXISTS, name);
            }
            mapFields.add(-1);
            log.debug("Headet col {} has index {} mapped with field {} at class {}", null, (mapFields.size() - 1), name, field.getTClass());
        }
        return mapFields;
    }
}
