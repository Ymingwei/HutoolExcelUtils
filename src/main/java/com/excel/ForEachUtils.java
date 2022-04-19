package com.excel;

import java.util.Objects;
import java.util.function.BiConsumer;

/**
 * for循环封装
 */
public class ForEachUtils {
    
    /**
     * 
     * @param <T>
     * @param startIndex 开始遍历的索引
     * @param elements 集合
     * @param action 
     */
    public static <T> void forEach(Integer startIndex,Iterable<? extends T> elements, BiConsumer<Integer, ? super T> action) {
        Objects.requireNonNull(elements);
        Objects.requireNonNull(action);
        startIndex = startIndex < 0 ? 0 : startIndex;
        Integer index = 0;
        for (T element : elements) {
            index++;
            if(index <= startIndex) {
                continue;
            }
            action.accept(index-1, element);
        }
    }
}