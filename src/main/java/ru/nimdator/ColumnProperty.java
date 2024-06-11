package ru.nimdator;

import lombok.Getter;
import lombok.RequiredArgsConstructor;
import lombok.ToString;

@Getter
@RequiredArgsConstructor
@ToString
public final class ColumnProperty {
    private final String columnName;
    private final int columnIndex;
    private final Class<?> tClass;

}
