package ru.nimdator;

import java.util.Objects;

public class ExcelPair<K, V> {
    /** Key. */
    private final K key;
    /** Value. */
    private final V value;

    /**
     * Create an entry representing a mapping from the specified key to the
     * specified value.
     *
     * @param k Key (first element of the pair).
     * @param v Value (second element of the pair).
     */
    public ExcelPair(K k, V v) {
        key = k;
        value = v;
    }

    /**
     * Create an entry representing the same mapping as the specified entry.
     *
     * @param entry Entry to copy.
     */
    public ExcelPair(ExcelPair<? extends K, ? extends V> entry) {
        this(entry.getKey(), entry.getValue());
    }

    /**
     * Get the key.
     *
     * @return the key (first element of the pair).
     */
    public K getKey() {
        return key;
    }

    /**
     * Get the value.
     *
     * @return the value (second element of the pair).
     */
    public V getValue() {
        return value;
    }

    /**
     * Compare the specified object with this entry for equality.
     *
     * @param o Object.
     * @return {@code true} if the given object is also a map entry and
     * the two entries represent the same mapping.
     */
    @Override
    public boolean equals(Object o) {
        if (this == o) {
            return true;
        }

        if (!(o instanceof ExcelPair<?, ?> oP)) {
            return false;
        }

        return (Objects.equals(key, oP.key)) &&
                    (Objects.equals(value, oP.value));
    }

    /**
     * Compute a hash code.
     *
     * @return the hash code value.
     */
    @Override
    public int hashCode() {
        int result = key == null ? 0 : key.hashCode();

        final int h = value == null ? 0 : value.hashCode();
        result = 37 * result + h ^ (h >>> 16);

        return result;
    }

    /** {@inheritDoc} */
    @Override
    public String toString() {
        return "[" + getKey() + ", " + getValue() + "]";
    }

    /**
     * Convenience factory method that calls the
     * {@link #ExcelPair(Object, Object) constructor}.
     *
     * @param <K> the key type
     * @param <V> the value type
     * @param k First element of the pair.
     * @param v Second element of the pair.
     * @return a new {@code Pair} containing {@code k} and {@code v}.
     * @since 3.3
     */
    public static <K, V> ExcelPair<K, V> create(K k, V v) {
        return new ExcelPair<K, V>(k, v);
    }
}
