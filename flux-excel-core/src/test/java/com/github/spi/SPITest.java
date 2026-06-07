package com.github.spi;

import java.util.ServiceLoader;

/**
 * @author Vico
 * @create 2023-05-17 11:22
 */
public class SPITest {
    public static void main(String[] args) {
        ServiceLoader<Search> serviceLoader = ServiceLoader.load(Search.class);
        for (Search search : serviceLoader) {
            search.searchDoc("hello word");
        }
    }
}
