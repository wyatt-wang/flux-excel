package com.github.spi;

import java.util.List;

/**
 * @author Vico
 * @create 2023-05-17 11:18
 */
public class DatabaseSearch implements Search{
    @Override
    public List<String> searchDoc(String keyword) {
        System.out.println("数据搜索 "+keyword);
        return null;
    }
}