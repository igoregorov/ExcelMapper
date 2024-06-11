package ru.nimdator;

import java.math.BigDecimal;
import java.util.Date;
import java.util.HashMap;
import java.util.Map;
import java.util.function.DoubleFunction;
import java.util.function.Function;

public abstract class CellData {
    private static final Function<String, String> cleanString = ConvertNumber::cleanString;
    private static final Map<Class<?>, Function<String, ?>> mapFromStringTo = new HashMap<>();
    static {
        mapFromStringTo.put(Long.class, cleanString.andThen(Long::parseLong));
        mapFromStringTo.put(Integer.class, cleanString.andThen(Integer::parseInt));
        mapFromStringTo.put(Double.class, cleanString.andThen(Double::parseDouble));
        mapFromStringTo.put(Short.class, cleanString.andThen(Short::parseShort));
        mapFromStringTo.put(BigDecimal.class, cleanString.andThen(ConvertNumber::toBigDecimal));
        mapFromStringTo.put(String.class, x -> x);
    }
    private static final Map<Class<?>, DoubleFunction<?>> mapFromDoubleTo = new HashMap<>();
    static {
        mapFromDoubleTo.put(Double.class, x -> x);
        mapFromDoubleTo.put(Integer.class, ConvertNumber::toInt);
        mapFromDoubleTo.put(Long.class, ConvertNumber::toLong);
        mapFromDoubleTo.put(Short.class, ConvertNumber::toShort);
        mapFromDoubleTo.put(BigDecimal.class, ConvertNumber::toBigDecimal);
        mapFromDoubleTo.put(String.class, ConvertNumber::toString);
    }

    public static <R> R getObjCellData(double obj, Class<R> requiredClass) {

        return requiredClass.cast(mapFromDoubleTo.get(requiredClass).apply(obj));
    }

    public static <R> R getObjCellData(String obj, Class<R> requiredClass) {
        return requiredClass.cast(mapFromStringTo.get(requiredClass).apply(obj));
    }

    public static <R> R getObjCellData(Date obj, Class<R> requiredClass) {
        return requiredClass.cast(obj);
    }

    public static <R> R getObjCellData(Class<R> requiredClass) {
        if (requiredClass.isAssignableFrom(Date.class)) return null;
        return requiredClass.cast(requiredClass.isAssignableFrom(String.class) ? "" : Double.NaN);
    }

    interface ConvertNumber {
        static BigDecimal toBigDecimal(Double a) {
            return BigDecimal.valueOf(a);
        }
        static short toShort(Double a) {
            return a.shortValue();
        }
        static int toInt(Double a) {
            return a.intValue();
        }
        static long toLong(Double a) {
            return a.longValue();
        }
        static String toString(Double a) {
            if (a % 1 == 0D) return String.valueOf(a.longValue());
            return String.valueOf(a);
        }
        static BigDecimal toBigDecimal(String s) {
            return BigDecimal.valueOf(Double.parseDouble(s));
        }

        static String cleanString(String obj) {
            return obj.trim().replace(',', '.');
        }
    }
}
